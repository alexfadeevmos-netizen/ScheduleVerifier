package com.example.scheduleverifier

import android.content.Context
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.text.PDFTextStripper
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook

enum class AuditVerdict(
    val code: Int,
    val badgeText: String,
    val textColor: Color,
    val bgColor: Color
) {
    CRITICAL_NOT_CANCELLED(
        1,
        "ОШИБКА: НЕ ОТМЕНЁН В EXCEL",
        Color(0xFFB71C1C),
        Color(0xFFFFEBEE)
    ),
    CRITICAL_NOT_ASSIGNED(
        2,
        "ОШИБКА: НЕ ВНЕСЁН В EXCEL",
        Color(0xFFC62828),
        Color(0xFFFFCDD2)
    ),
    ERROR_TIME_OR_STOP(
        3,
        "НЕСООТВЕТСТВИЕ ТЕЛЕГРАММЕ",
        Color(0xFFE65100),
        Color(0xFFFFF3E0)
    ),
    NEEDS_MANUAL_CHECK(
        4,
        "ПРОВЕРИТЬ УЧАСТОК / ОСТАНОВКИ",
        Color(0xFFF57F17),
        Color(0xFFFFFDE7)
    ),
    EXECUTED_CORRECTLY(
        5,
        "ТЕЛЕГРАММА ИСПОЛНЕНА",
        Color(0xFF1B5E20),
        Color(0xFFE8F5E9)
    ),
    UNCHANGED_BASE_SCHEDULE(
        6,
        "БЕЗ ИЗМЕНЕНИЙ (ПО ГРАФИКУ)",
        Color(0xFF455A64),
        Color(0xFFECEFF1)
    )
}

enum class TelegramDirective(val title: String) {
    FULL_CANCEL("Полная отмена поезда"),
    PARTIAL_CANCEL("Укорочение маршрута / отмена на участке"),
    ASSIGN_TRAIN("Назначение поезда"),
    SKIP_STOPS("Проследование без остановки"),
    TIME_CHANGE("Изменение времени (расписания)"),
    INFO_MENTION("Упоминание в телеграмме")
}

data class TelegramOrderBlock(
    val pageNumber: Int,
    val blockIndex: Int,
    val locationLabel: String,
    val rawNumbers: Set<String>,
    val displayNumber: String,
    val directive: TelegramDirective,
    val requiredTimes: List<String>,
    val mentionedStops: List<String>,
    val fullText: String
)

data class ScheduleRowRecord(
    val sheetName: String,
    val rowNumber: Int,
    val locationLabel: String,
    val rawNumbers: Set<String>,
    val displayNumber: String,
    val actualTimes: List<String>,
    val isMarkedCancelled: Boolean,
    val fullText: String
)

data class TrainAuditItem(
    val trainDisplay: String,
    val verdict: AuditVerdict,
    val directiveTitle: String,
    val telegramLocations: List<String>,
    val scheduleLocations: List<String>,
    val telegramTargetTimes: List<String>,
    val scheduleActualTimes: List<String>,
    val discrepancySummary: String,
    val requiredAction: String,
    val telegramQuotes: List<String>,
    val scheduleQuotes: List<String>
)

data class AuditReport(
    val telegramFileName: String,
    val scheduleFileName: String,
    val telegramOrdersCount: Int,
    val scheduleRowsCount: Int,
    val items: List<TrainAuditItem>
)

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        PDFBoxResourceLoader.init(applicationContext)

        setContent {
            MaterialTheme(
                colorScheme = lightColorScheme(
                    primary = Color(0xFF0D47A1),
                    secondary = Color(0xFF1565C0),
                    background = Color(0xFFF4F6F9)
                )
            ) {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    TelegramPriorityVerifierScreen()
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun TelegramPriorityVerifierScreen() {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()

    var telegramUri by remember { mutableStateOf<Uri?>(null) }
    var telegramName by remember { mutableStateOf("Файл телеграммы не выбран") }

    var scheduleUri by remember { mutableStateOf<Uri?>(null) }
    var scheduleName by remember { mutableStateOf("Файл расписания не выбран") }

    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var report by remember { mutableStateOf<AuditReport?>(null) }

    var selectedTab by remember { mutableIntStateOf(0) }
    var searchQuery by remember { mutableStateOf("") }

    val telegramPicker = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let {
            telegramUri = it
            telegramName = getFileName(context, it)
        }
    }

    val schedulePicker = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let {
            scheduleUri = it
            scheduleName = getFileName(context, it)
        }
    }

    val exportExcelLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ) { uri ->
        uri?.let { targetUri ->
            val currentReport = report ?: return@let
            scope.launch {
                try {
                    exportReportToExcel(context, targetUri, currentReport)
                    Toast.makeText(context, "Отчёт успешно сохранён в Excel!", Toast.LENGTH_LONG).show()
                } catch (e: Exception) {
                    Toast.makeText(context, "Ошибка сохранения: ${e.localizedMessage}", Toast.LENGTH_LONG).show()
                }
            }
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            "Ревизор Расписаний (Приоритет: Телеграмма)",
                            fontWeight = FontWeight.Bold,
                            fontSize = 17.sp
                        )
                        Text(
                            "Эталон: Телеграмма (PDF/Excel) ➔ Проверка: График (Excel/PDF)",
                            fontSize = 11.sp,
                            color = Color.White.copy(alpha = 0.85f)
                        )
                    }
                },
                actions = {
                    if (report != null) {
                        IconButton(
                            onClick = {
                                exportExcelLauncher.launch("Сверка_по_телеграмме.xlsx")
                            }
                        ) {
                            Icon(
                                Icons.Default.Download,
                                contentDescription = "Выгрузить отчёт в Excel",
                                tint = Color.White
                            )
                        }
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.primary,
                    titleContentColor = Color.White
                )
            )
        }
    ) { padding ->
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .padding(12.dp)
        ) {
            Card(
                modifier = Modifier.fillMaxWidth(),
                shape = RoundedCornerShape(12.dp),
                colors = CardDefaults.cardColors(containerColor = Color.White),
                elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
            ) {
                Column(
                    modifier = Modifier.padding(12.dp),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    RoleFilePickerRow(
                        badge = "ПРИОРИТЕТ 1 • ЭТАЛОН",
                        badgeColor = Color(0xFFB71C1C),
                        label = "Оперативная телеграмма (PDF или Excel):",
                        fileName = telegramName,
                        isSelected = telegramUri != null,
                        onSelectClick = { telegramPicker.launch(arrayOf("*/*")) }
                    )

                    HorizontalDivider(color = Color(0xFFEEEEEE))

                    RoleFilePickerRow(
                        badge = "ПРИОРИТЕТ 2 • ПРОВЕРЯЕМЫЙ ГРАФИК",
                        badgeColor = Color(0xFF1565C0),
                        label = "Фактическое расписание (Excel или PDF):",
                        fileName = scheduleName,
                        isSelected = scheduleUri != null,
                        onSelectClick = { schedulePicker.launch(arrayOf("*/*")) }
                    )

                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        Button(
                            onClick = {
                                if (telegramUri != null && scheduleUri != null) {
                                    isLoading = true
                                    errorMessage = null
                                    scope.launch {
                                        try {
                                            report = runTelegramPriorityAudit(
                                                context,
                                                telegramUri!!,
                                                telegramName,
                                                scheduleUri!!,
                                                scheduleName
                                            )
                                            selectedTab = 0
                                        } catch (e: Exception) {
                                            errorMessage = "Ошибка анализа: ${e.localizedMessage ?: e.javaClass.simpleName}"
                                        } finally {
                                            isLoading = false
                                        }
                                    }
                                }
                            },
                            enabled = telegramUri != null && scheduleUri != null && !isLoading,
                            modifier = Modifier.weight(1f),
                            shape = RoundedCornerShape(8.dp)
                        ) {
                            if (isLoading) {
                                CircularProgressIndicator(
                                    modifier = Modifier.size(18.dp),
                                    color = Color.White,
                                    strokeWidth = 2.dp
                                )
                                Spacer(modifier = Modifier.width(8.dp))
                                Text("Сверка с телеграммой...")
                            } else {
                                Icon(Icons.Default.FactCheck, contentDescription = null)
                                Spacer(modifier = Modifier.width(8.dp))
                                Text("Проверить исполнение телеграммы")
                            }
                        }

                        if (report != null) {
                            OutlinedButton(
                                onClick = { exportExcelLauncher.launch("Сверка_по_телеграмме.xlsx") },
                                shape = RoundedCornerShape(8.dp)
                            ) {
                                Icon(Icons.Default.TableView, contentDescription = null, modifier = Modifier.size(18.dp))
                                Spacer(modifier = Modifier.width(4.dp))
                                Text("В Excel")
                            }
                        }
                    }
                }
            }

            errorMessage?.let { msg ->
                Spacer(modifier = Modifier.height(8.dp))
                Card(
                    colors = CardDefaults.cardColors(containerColor = Color(0xFFFFEBEE)),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text(
                        text = msg,
                        color = Color(0xFFB71C1C),
                        modifier = Modifier.padding(12.dp),
                        fontSize = 13.sp
                    )
                }
            }

            report?.let { currentReport ->
                Spacer(modifier = Modifier.height(10.dp))

                val allItems = currentReport.items
                val violations = allItems.filter {
                    it.verdict == AuditVerdict.CRITICAL_NOT_CANCELLED ||
                    it.verdict == AuditVerdict.CRITICAL_NOT_ASSIGNED ||
                    it.verdict == AuditVerdict.ERROR_TIME_OR_STOP ||
                    it.verdict == AuditVerdict.NEEDS_MANUAL_CHECK
                }
                val executed = allItems.filter { it.verdict == AuditVerdict.EXECUTED_CORRECTLY }
                val untouched = allItems.filter { it.verdict == AuditVerdict.UNCHANGED_BASE_SCHEDULE }

                ScrollableTabRow(
                    selectedTabIndex = selectedTab,
                    edgePadding = 0.dp,
                    containerColor = Color.White,
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Tab(
                        selected = selectedTab == 0,
                        onClick = { selectedTab = 0 },
                        text = {
                            Text(
                                "Ошибки в Excel (${violations.size})",
                                color = Color(0xFFC62828),
                                fontWeight = FontWeight.Bold
                            )
                        }
                    )
                    Tab(
                        selected = selectedTab == 1,
                        onClick = { selectedTab = 1 },
                        text = {
                            Text(
                                "Исполнено точно (${executed.size})",
                                color = Color(0xFF2E7D32)
                            )
                        }
                    )
                    Tab(
                        selected = selectedTab == 2,
                        onClick = { selectedTab = 2 },
                        text = { Text("Все из телеграммы (${violations.size + executed.size})") }
                    )
                    Tab(
                        selected = selectedTab == 3,
                        onClick = { selectedTab = 3 },
                        text = { Text("Остальные в Excel (${untouched.size})") }
                    )
                }

                Spacer(modifier = Modifier.height(8.dp))

                OutlinedTextField(
                    value = searchQuery,
                    onValueChange = { searchQuery = it },
                    placeholder = { Text("Фильтр по № поезда (напр. 6502, 7104), станции или стр. PDF...") },
                    leadingIcon = { Icon(Icons.Default.Search, contentDescription = null) },
                    trailingIcon = {
                        if (searchQuery.isNotEmpty()) {
                            IconButton(onClick = { searchQuery = "" }) {
                                Icon(Icons.Default.Clear, contentDescription = null)
                            }
                        }
                    },
                    singleLine = true,
                    modifier = Modifier.fillMaxWidth(),
                    colors = OutlinedTextFieldDefaults.colors(
                        focusedContainerColor = Color.White,
                        unfocusedContainerColor = Color.White
                    )
                )

                Spacer(modifier = Modifier.height(8.dp))

                val baseList = when (selectedTab) {
                    0 -> violations
                    1 -> executed
                    2 -> violations + executed
                    else -> untouched
                }

                val filteredList = baseList.filter { item ->
                    searchQuery.isBlank() ||
                    item.trainDisplay.contains(searchQuery, ignoreCase = true) ||
                    item.discrepancySummary.contains(searchQuery, ignoreCase = true) ||
                    item.telegramLocations.any { it.contains(searchQuery, ignoreCase = true) } ||
                    item.telegramQuotes.any { it.contains(searchQuery, ignoreCase = true) } ||
                    item.scheduleQuotes.any { it.contains(searchQuery, ignoreCase = true) }
                }

                LazyColumn(
                    modifier = Modifier.fillMaxSize(),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    items(filteredList) { item ->
                        AuditResultCard(item)
                    }
                }
            }
        }
    }
}

@Composable
fun RoleFilePickerRow(
    badge: String,
    badgeColor: Color,
    label: String,
    fileName: String,
    isSelected: Boolean,
    onSelectClick: () -> Unit
) {
    Row(
        modifier = Modifier.fillMaxWidth(),
        verticalAlignment = Alignment.CenterVertically,
        horizontalArrangement = Arrangement.SpaceBetween
    ) {
        Column(modifier = Modifier.weight(1f)) {
            Surface(
                color = badgeColor.copy(alpha = 0.12f),
                shape = RoundedCornerShape(4.dp)
            ) {
                Text(
                    text = badge,
                    color = badgeColor,
                    fontSize = 10.sp,
                    fontWeight = FontWeight.Bold,
                    modifier = Modifier.padding(horizontal = 6.dp, vertical = 2.dp)
                )
            }
            Spacer(modifier = Modifier.height(2.dp))
            Text(label, fontSize = 12.sp, color = Color.Gray)
            Text(
                text = fileName,
                fontSize = 14.sp,
                fontWeight = if (isSelected) FontWeight.Bold else FontWeight.Normal,
                color = if (isSelected) Color(0xFF0D47A1) else Color.DarkGray
            )
        }
        OutlinedButton(onClick = onSelectClick) {
            Icon(Icons.Default.UploadFile, contentDescription = null, modifier = Modifier.size(18.dp))
            Spacer(modifier = Modifier.width(4.dp))
            Text("Выбрать")
        }
    }
}

@Composable
fun AuditResultCard(item: TrainAuditItem) {
    var expanded by remember { mutableStateOf(false) }

    Card(
        modifier = Modifier.fillMaxWidth(),
        shape = RoundedCornerShape(10.dp),
        colors = CardDefaults.cardColors(containerColor = Color.White),
        elevation = CardDefaults.cardElevation(defaultElevation = 1.dp)
    ) {
        Column(modifier = Modifier.padding(12.dp)) {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Icon(
                        imageVector = Icons.Default.Train,
                        contentDescription = null,
                        tint = item.verdict.textColor,
                        modifier = Modifier.size(22.dp)
                    )
                    Spacer(modifier = Modifier.width(6.dp))
                    Column {
                        Text(
                            text = "Поезд № ${item.trainDisplay}",
                            fontWeight = FontWeight.Bold,
                            fontSize = 16.sp
                        )
                        Text(
                            text = "Приказ в телеграмме: ${item.directiveTitle}",
                            fontSize = 11.sp,
                            color = Color(0xFF546E7A)
                        )
                    }
                }

                Surface(
                    color = item.verdict.bgColor,
                    shape = RoundedCornerShape(6.dp)
                ) {
                    Text(
                        text = item.verdict.badgeText,
                        color = item.verdict.textColor,
                        fontWeight = FontWeight.Bold,
                        fontSize = 11.sp,
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 4.dp)
                    )
                }
            }

            Spacer(modifier = Modifier.height(8.dp))

            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                if (item.telegramLocations.isNotEmpty()) {
                    Text(
                        text = "📌 Телеграмма: ${item.telegramLocations.joinToString(", ")}",
                        fontSize = 12.sp,
                        fontWeight = FontWeight.SemiBold,
                        color = Color(0xFFB71C1C)
                    )
                }
                if (item.scheduleLocations.isNotEmpty()) {
                    Text(
                        text = "📊 Расписание: ${item.scheduleLocations.joinToString(", ")}",
                        fontSize = 12.sp,
                        fontWeight = FontWeight.SemiBold,
                        color = Color(0xFF1565C0)
                    )
                } else {
                    Text(
                        text = "📊 Расписание: отсутствует в таблице",
                        fontSize = 12.sp,
                        fontWeight = FontWeight.SemiBold,
                        color = Color.Gray
                    )
                }
            }

            Spacer(modifier = Modifier.height(6.dp))

            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .background(Color(0xFFF8F9FA), RoundedCornerShape(6.dp))
                    .padding(8.dp),
                verticalArrangement = Arrangement.spacedBy(4.dp)
            ) {
                Text(
                    text = item.discrepancySummary,
                    fontSize = 13.sp,
                    color = Color(0xFF212121)
                )
                if (item.requiredAction.isNotEmpty()) {
                    Text(
                        text = "👉 Что исправить в Excel: ${item.requiredAction}",
                        fontSize = 12.sp,
                        fontWeight = FontWeight.Bold,
                        color = item.verdict.textColor
                    )
                }
            }

            if (item.telegramTargetTimes.isNotEmpty() || item.scheduleActualTimes.isNotEmpty()) {
                Spacer(modifier = Modifier.height(6.dp))
                Text(
                    text = "Время по Телеграмме (Эталон): ${item.telegramTargetTimes.ifEmpty { listOf("—") }.joinToString(", ")}\n" +
                           "Время в Расписании (Факт): ${item.scheduleActualTimes.ifEmpty { listOf("—") }.joinToString(", ")}",
                    fontSize = 12.sp,
                    color = Color(0xFF37474F),
                    fontWeight = FontWeight.Medium
                )
            }

            TextButton(
                onClick = { expanded = !expanded },
                contentPadding = PaddingValues(0.dp)
            ) {
                Text(
                    text = if (expanded) "Скрыть цитаты из документов ▲" else "Показать точный текст из Телеграммы и Расписания ▼",
                    fontSize = 12.sp
                )
            }

            if (expanded) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .background(Color(0xFFECEFF1), RoundedCornerShape(6.dp))
                        .padding(8.dp),
                    verticalArrangement = Arrangement.spacedBy(6.dp)
                ) {
                    if (item.telegramQuotes.isNotEmpty()) {
                        Text(
                            "Текст приказа в Телеграмме (Эталон):",
                            fontWeight = FontWeight.Bold,
                            fontSize = 12.sp,
                            color = Color(0xFFB71C1C)
                        )
                        item.telegramQuotes.forEach { quote ->
                            Text("• $quote", fontSize = 12.sp)
                        }
                    }
                    if (item.scheduleQuotes.isNotEmpty()) {
                        Text(
                            "Строки в Расписании Excel (Факт):",
                            fontWeight = FontWeight.Bold,
                            fontSize = 12.sp,
                            color = Color(0xFF0D47A1)
                        )
                        item.scheduleQuotes.forEach { quote ->
                            Text("• $quote", fontSize = 12.sp)
                        }
                    }
                }
            }
        }
    }
}

private val COMPOSITE_TRAIN_REGEX = Regex("""(?<!\d)([67]\d{3}(?:/[67]\d{3})?)(?!\d)""")
private val SINGLE_TRAIN_REGEX = Regex("""(?<!\d)[67]\d{3}(?!\d)""")
private val TIME_REGEX = Regex("""\b([01]?\d|2[0-3])[:.-]([0-5]\d)\b""")
private val STATION_HINT_REGEX = Regex(
    """(?:ст\.|о\.п\.|пл\.|участке|сообщением|маршрутом)\s+([А-ЯЁ][а-яёА-ЯЁ0-9\-\s]{2,25})""",
    RegexOption.IGNORE_CASE
)

suspend fun runTelegramPriorityAudit(
    context: Context,
    telegramUri: Uri,
    telegramName: String,
    scheduleUri: Uri,
    scheduleName: String
): AuditReport = withContext(Dispatchers.IO) {

    val telegramOrders = parseTelegramOrders(context, telegramUri, telegramName)
    val scheduleRows = parseScheduleRows(context, scheduleUri, scheduleName)

    val usedScheduleRows = mutableSetOf<ScheduleRowRecord>()
    val auditItems = mutableListOf<TrainAuditItem>()

    val groupedOrders = mutableMapOf<String, MutableList<TelegramOrderBlock>>()
    for (order in telegramOrders) {
        val canonicalKey = order.rawNumbers.sorted().joinToString("/")
        groupedOrders.getOrPut(canonicalKey) { mutableListOf() }.add(order)
    }

    for ((_, ordersForTrain) in groupedOrders) {
        val allNumbersOfTrain = ordersForTrain.flatMap { it.rawNumbers }.toSet()
        val displayTrain = ordersForTrain.first().displayNumber

        val matchingExcelRows = scheduleRows.filter { row ->
            row.rawNumbers.intersect(allNumbersOfTrain).isNotEmpty()
        }
        usedScheduleRows.addAll(matchingExcelRows)

        val primaryDirective = when {
            ordersForTrain.any { it.directive == TelegramDirective.FULL_CANCEL } -> TelegramDirective.FULL_CANCEL
            ordersForTrain.any { it.directive == TelegramDirective.PARTIAL_CANCEL } -> TelegramDirective.PARTIAL_CANCEL
            ordersForTrain.any { it.directive == TelegramDirective.ASSIGN_TRAIN } -> TelegramDirective.ASSIGN_TRAIN
            ordersForTrain.any { it.directive == TelegramDirective.SKIP_STOPS } -> TelegramDirective.SKIP_STOPS
            ordersForTrain.any { it.directive == TelegramDirective.TIME_CHANGE } -> TelegramDirective.TIME_CHANGE
            else -> TelegramDirective.INFO_MENTION
        }

        val tgLocations = ordersForTrain.map { it.locationLabel }.distinct()
        val exLocations = matchingExcelRows.map { it.locationLabel }.distinct()

        val tgTimes = ordersForTrain.flatMap { it.requiredTimes }.distinct()
        val exTimes = matchingExcelRows.flatMap { it.actualTimes }.distinct()

        val tgQuotes = ordersForTrain.map { "[${it.locationLabel}] ${it.fullText}" }
        val exQuotes = matchingExcelRows.map { "[${it.locationLabel}] ${it.fullText}" }

        val mentionedStops = ordersForTrain.flatMap { it.mentionedStops }.distinct()

        val (verdict, summary, action) = when (primaryDirective) {

            TelegramDirective.FULL_CANCEL -> {
                val activeRowsInExcel = matchingExcelRows.filter { !it.isMarkedCancelled }
                if (activeRowsInExcel.isEmpty()) {
                    Triple(
                        AuditVerdict.EXECUTED_CORRECTLY,
                        "✅ По телеграмме (${tgLocations.joinToString()}) поезд ОТМЕНЁН. В действующем расписании Excel он корректно отсутствует (или помечен отменённым).",
                        ""
                    )
                } else {
                    Triple(
                        AuditVerdict.CRITICAL_NOT_CANCELLED,
                        "❌ КРИТИЧЕСКАЯ ОШИБКА: Согласно телеграмме (${tgLocations.joinToString()}) поезд № $displayTrain ОТМЕНЯЕТСЯ, но в файле расписания (${exLocations.joinToString()}) он оставлен в графике!",
                        "Удалить поезд № $displayTrain из расписания Excel (${exLocations.joinToString()}) или поставить отметку об отмене."
                    )
                }
            }

            TelegramDirective.ASSIGN_TRAIN -> {
                if (matchingExcelRows.isEmpty()) {
                    Triple(
                        AuditVerdict.CRITICAL_NOT_ASSIGNED,
                        "❌ КРИТИЧЕСКАЯ ОШИБКА: Согласно телеграмме (${tgLocations.joinToString()}) НАЗНАЧАЕТСЯ поезд № $displayTrain, но в файле расписания Excel он полностью отсутствует!",
                        "Внести поезд № $displayTrain в расписание Excel со временем из телеграммы: ${tgTimes.joinToString(", ").ifEmpty { "см. текст телеграммы" }}."
                    )
                } else {
                    val missingTimes = tgTimes.filter { it !in exTimes }
                    if (missingTimes.isEmpty()) {
                        Triple(
                            AuditVerdict.EXECUTED_CORRECTLY,
                            "✅ Поезд № $displayTrain назначен по телеграмме (${tgLocations.joinToString()}) и корректно внесён в Excel (${exLocations.joinToString()}) с точным совпадением времени.",
                            ""
                        )
                    } else {
                        Triple(
                            AuditVerdict.ERROR_TIME_OR_STOP,
                            "⚠️ Поезд № $displayTrain внесён в Excel (${exLocations.joinToString()}), но время в таблице НЕ совпадает с телеграммой (${tgLocations.joinToString()}). Отсутствуют контрольные точки телеграммы: ${missingTimes.joinToString(", ")}.",
                            "Заменить время в Excel (${exLocations.joinToString()}) на эталонное из телеграммы: ${tgTimes.joinToString(", ")}."
                        )
                    }
                }
            }

            TelegramDirective.PARTIAL_CANCEL, TelegramDirective.SKIP_STOPS, TelegramDirective.TIME_CHANGE, TelegramDirective.INFO_MENTION -> {
                if (matchingExcelRows.isEmpty()) {
                    Triple(
                        AuditVerdict.CRITICAL_NOT_ASSIGNED,
                        "❌ ОШИБКА: В телеграмме (${tgLocations.joinToString()}) изменяется расписание поезда № $displayTrain, однако в файле Excel этот поезд вообще не найден!",
                        "Проверить наличие поезда № $displayTrain в таблице Excel и внести изменения согласно ${tgLocations.joinToString()}."
                    )
                } else if (tgTimes.isNotEmpty()) {
                    val missingTargetTimes = tgTimes.filter { it !in exTimes }
                    if (missingTargetTimes.isEmpty()) {
                        Triple(
                            AuditVerdict.EXECUTED_CORRECTLY,
                            "✅ Все контрольные точки времени из телеграммы (${tgTimes.joinToString()}) точно внесены в расписание Excel (${exLocations.joinToString()}).",
                            ""
                        )
                    } else {
                        val stopsNote = if (mentionedStops.isNotEmpty()) " (участок/станции: ${mentionedStops.joinToString()})" else ""
                        Triple(
                            AuditVerdict.ERROR_TIME_OR_STOP,
                            "⚠️ РАСХОЖДЕНИЕ С ТЕЛЕГРАММОЙ$stopsNote: На ${tgLocations.joinToString()} установлено время [${tgTimes.joinToString()}], но в Excel (${exLocations.joinToString()}) указано [${exTimes.joinToString()}]. Не исполнены точки: ${missingTargetTimes.joinToString()}.",
                            "Исправить в Excel (${exLocations.joinToString()}) время на эталонное из телеграммы: ${missingTargetTimes.joinToString()}."
                        )
                    }
                } else {
                    val stopsInfo = if (mentionedStops.isNotEmpty()) " (${mentionedStops.joinToString()})" else ""
                    Triple(
                        AuditVerdict.NEEDS_MANUAL_CHECK,
                        "ℹ️ В телеграмме (${tgLocations.joinToString()}) указан приказ: «${primaryDirective.title}»$stopsInfo без явных меток ЧЧ:ММ. Поезд найден в Excel на ${exLocations.joinToString()}.",
                        "Сверить остановки и маршрут в строке ${exLocations.joinToString()} с текстом телеграммы на ${tgLocations.joinToString()}."
                    )
                }
            }
        }

        auditItems.add(
            TrainAuditItem(
                trainDisplay = displayTrain,
                verdict = verdict,
                directiveTitle = primaryDirective.title,
                telegramLocations = tgLocations,
                scheduleLocations = exLocations,
                telegramTargetTimes = tgTimes,
                scheduleActualTimes = exTimes,
                discrepancySummary = summary,
                requiredAction = action,
                telegramQuotes = tgQuotes,
                scheduleQuotes = exQuotes
            )
        )
    }

    val untouchedRows = scheduleRows.filter { it !in usedScheduleRows }
    val untouchedByTrain = untouchedRows.groupBy { it.displayNumber }
    for ((trainDisp, rows) in untouchedByTrain) {
        auditItems.add(
            TrainAuditItem(
                trainDisplay = trainDisp,
                verdict = AuditVerdict.UNCHANGED_BASE_SCHEDULE,
                directiveTitle = "Не упоминается в телеграмме",
                telegramLocations = emptyList(),
                scheduleLocations = rows.map { it.locationLabel }.distinct(),
                telegramTargetTimes = emptyList(),
                scheduleActualTimes = rows.flatMap { it.actualTimes }.distinct(),
                discrepancySummary = "Поезд следует базовым расписанием в Excel. В оперативной телеграмме изменений по нему нет.",
                requiredAction = "",
                telegramQuotes = emptyList(),
                scheduleQuotes = rows.map { "[${it.locationLabel}] ${it.fullText}" }
            )
        )
    }

    val sortedItems = auditItems.sortedWith(
        compareBy<TrainAuditItem> { it.verdict.code }.thenBy { it.trainDisplay }
    )

    AuditReport(
        telegramFileName = telegramName,
        scheduleFileName = scheduleName,
        telegramOrdersCount = groupedOrders.size,
        scheduleRowsCount = scheduleRows.size,
        items = sortedItems
    )
}

fun parseTelegramOrders(context: Context, uri: Uri, fileName: String): List<TelegramOrderBlock> {
    val lower = fileName.lowercase()
    val rawLinesWithPage = mutableListOf<Pair<Int, String>>()

    if (lower.endsWith(".pdf")) {
        context.contentResolver.openInputStream(uri)?.use { inputStream ->
            PDDocument.load(inputStream).use { doc ->
                val stripper = PDFTextStripper()
                for (page in 1..doc.numberOfPages) {
                    stripper.startPage = page
                    stripper.endPage = page
                    val pageText = stripper.getText(doc)
                    pageText.lines()
                        .map { it.trim() }
                        .filter { it.isNotEmpty() }
                        .forEach { line -> rawLinesWithPage.add(page to line) }
                }
            }
        }
    } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        val formatter = DataFormatter()
        context.contentResolver.openInputStream(uri)?.use { inputStream ->
            WorkbookFactory.create(inputStream).use { wb ->
                for (s in 0 until wb.numberOfSheets) {
                    val sheet = wb.getSheetAt(s)
                    for (row in sheet) {
                        val rowNum = row.rowNum + 1
                        val text = row.mapNotNull { formatter.formatCellValue(it).trim().takeIf { v -> v.isNotEmpty() } }
                            .joinToString(" ")
                        if (text.isNotEmpty()) {
                            rawLinesWithPage.add(rowNum to text)
                        }
                    }
                }
            }
        }
    } else {
        throw IllegalArgumentException("Формат файла телеграммы '$fileName' не поддерживается.")
    }

    val blocks = mutableListOf<TelegramOrderBlock>()
    var currentPage = 1
    var currentTrainsDisplay = ""
    var currentRawTrains = setOf<String>()
    val currentBuffer = mutableListOf<String>()
    var blockCounter = 0
    var inheritedDirective: TelegramDirective? = null

    fun flushCurrentBlock() {
        if (currentRawTrains.isNotEmpty() && currentBuffer.isNotEmpty()) {
            blockCounter++
            val joinedText = currentBuffer.joinToString(" ")
            val detectedDirective = detectTelegramDirective(joinedText) ?: inheritedDirective ?: TelegramDirective.INFO_MENTION
            val times = extractNormalizedTimes(joinedText)
            val stops = STATION_HINT_REGEX.findAll(joinedText).map { it.groupValues[1].trim() }.toList()

            blocks.add(
                TelegramOrderBlock(
                    pageNumber = currentPage,
                    blockIndex = blockCounter,
                    locationLabel = if (lower.endsWith(".pdf")) "Стр. $currentPage (блок $blockCounter)" else "Строка $currentPage",
                    rawNumbers = currentRawTrains,
                    displayNumber = currentTrainsDisplay,
                    directive = detectedDirective,
                    requiredTimes = times,
                    mentionedStops = stops,
                    fullText = joinedText
                )
            )
        }
        currentBuffer.clear()
        currentRawTrains = emptySet()
        currentTrainsDisplay = ""
    }

    for ((page, line) in rawLinesWithPage) {
        val compositeMatches = COMPOSITE_TRAIN_REGEX.findAll(line).map { it.value }.toList()
        val singleMatches = SINGLE_TRAIN_REGEX.findAll(line).map { it.value }.toSet()

        val headerDir = detectTelegramDirective(line)
        if (singleMatches.isEmpty() && headerDir != null && line.length < 120) {
            inheritedDirective = headerDir
        }

        if (singleMatches.isNotEmpty()) {
            if (currentRawTrains.isNotEmpty() && currentRawTrains != singleMatches) {
                flushCurrentBlock()
            }
            currentPage = page
            currentRawTrains = singleMatches
            currentTrainsDisplay = compositeMatches.firstOrNull() ?: singleMatches.joinToString("/")
            currentBuffer.add(line)
        } else if (currentRawTrains.isNotEmpty()) {
            currentBuffer.add(line)
        }
    }
    flushCurrentBlock()

    return blocks
}

fun parseScheduleRows(context: Context, uri: Uri, fileName: String): List<ScheduleRowRecord> {
    val lower = fileName.lowercase()
    val rows = mutableListOf<ScheduleRowRecord>()

    if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        val formatter = DataFormatter()
        context.contentResolver.openInputStream(uri)?.use { inputStream ->
            WorkbookFactory.create(inputStream).use { workbook ->
                for (s in 0 until workbook.numberOfSheets) {
                    val sheet = workbook.getSheetAt(s)
                    val sheetName = sheet.sheetName

                    for (row in sheet) {
                        val rowNum = row.rowNum + 1
                        val cells = mutableListOf<String>()
                        for (cell in row) {
                            val value = formatter.formatCellValue(cell).trim()
                            if (value.isNotEmpty()) cells.add(value)
                        }
                        if (cells.isNotEmpty()) {
                            val fullRowText = cells.joinToString(" | ")
                            val singleTrains = SINGLE_TRAIN_REGEX.findAll(fullRowText).map { it.value }.toSet()
                            val compositeDisp = COMPOSITE_TRAIN_REGEX.findAll(fullRowText).map { it.value }.firstOrNull()
                                ?: singleTrains.joinToString("/")
                            val times = extractNormalizedTimes(fullRowText)

                            if (singleTrains.isNotEmpty()) {
                                rows.add(
                                    ScheduleRowRecord(
                                        sheetName = sheetName,
                                        rowNumber = rowNum,
                                        locationLabel = "Лист '$sheetName', стр. $rowNum",
                                        rawNumbers = singleTrains,
                                        displayNumber = compositeDisp,
                                        actualTimes = times,
                                        isMarkedCancelled = fullRowText.lowercase().contains("отмен"),
                                        fullText = fullRowText
                                    )
                                )
                            }
                        }
                    }
                }
            }
        }
    } else if (lower.endsWith(".pdf")) {
        context.contentResolver.openInputStream(uri)?.use { inputStream ->
            PDDocument.load(inputStream).use { doc ->
                val stripper = PDFTextStripper()
                for (page in 1..doc.numberOfPages) {
                    stripper.startPage = page
                    stripper.endPage = page
                    stripper.getText(doc).lines().forEachIndexed { idx, line ->
                        val trimmed = line.trim()
                        val singleTrains = SINGLE_TRAIN_REGEX.findAll(trimmed).map { it.value }.toSet()
                        if (singleTrains.isNotEmpty()) {
                            rows.add(
                                ScheduleRowRecord(
                                    sheetName = "PDF",
                                    rowNumber = idx + 1,
                                    locationLabel = "Стр. $page, строка ${idx + 1}",
                                    rawNumbers = singleTrains,
                                    displayNumber = singleTrains.joinToString("/"),
                                    actualTimes = extractNormalizedTimes(trimmed),
                                    isMarkedCancelled = trimmed.lowercase().contains("отмен"),
                                    fullText = trimmed
                                )
                            )
                        }
                    }
                }
            }
        }
    } else {
        throw IllegalArgumentException("Формат файла расписания '$fileName' не поддерживается.")
    }

    return rows
}

fun detectTelegramDirective(text: String): TelegramDirective? {
    val lower = text.lowercase()
    return when {
        (lower.contains("на участке") && lower.contains("отмен")) ||
        lower.contains("укорач") ||
        lower.contains("сокращ") -> TelegramDirective.PARTIAL_CANCEL

        lower.contains("отменяется") ||
        lower.contains("отменяются") ||
        lower.contains("отменить") ||
        lower.contains("выводится из расписания") -> TelegramDirective.FULL_CANCEL

        lower.contains("назначается") ||
        lower.contains("назначаются") ||
        lower.contains("назначить") -> TelegramDirective.ASSIGN_TRAIN

        lower.contains("без остановк") ||
        lower.contains("без остановок") ||
        lower.contains("проследует") -> TelegramDirective.SKIP_STOPS

        lower.contains("изменяется") ||
        lower.contains("изменить") ||
        lower.contains("разработанным расписанием") ||
        lower.contains("измененным расписанием") -> TelegramDirective.TIME_CHANGE

        else -> null
    }
}

fun extractNormalizedTimes(text: String): List<String> {
    return TIME_REGEX.findAll(text)
        .map { match ->
            val h = match.groupValues[1].padStart(2, '0')
            val m = match.groupValues[2].padStart(2, '0')
            "$h:$m"
        }
        .distinct()
        .toList()
}

suspend fun exportReportToExcel(
    context: Context,
    targetUri: Uri,
    report: AuditReport
) = withContext(Dispatchers.IO) {
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Сверка по Телеграмме")

    val headerStyle = workbook.createCellStyle().apply {
        fillForegroundColor = IndexedColors.DARK_BLUE.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
        val font = workbook.createFont().apply {
            bold = true
            color = IndexedColors.WHITE.index
        }
        setFont(font)
    }

    val errorStyle = workbook.createCellStyle().apply {
        fillForegroundColor = IndexedColors.ROSE.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
    }

    val warnStyle = workbook.createCellStyle().apply {
        fillForegroundColor = IndexedColors.LIGHT_YELLOW.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
    }

    val okStyle = workbook.createCellStyle().apply {
        fillForegroundColor = IndexedColors.LIGHT_GREEN.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
    }

    val columns = listOf(
        "№ Поезда",
        "Статус (Приоритет: Телеграмма)",
        "Приказ в Телеграмме",
        "Стр. в Телеграмме (PDF)",
        "Строка в Расписании (Excel)",
        "Время по Телеграмме (Эталон)",
        "Время в Расписании (Факт)",
        "Заключение ревизора",
        "Что исправить в Excel"
    )

    val headerRow = sheet.createRow(0)
    columns.forEachIndexed { idx, title ->
        val cell = headerRow.createCell(idx)
        cell.setCellValue(title)
        cell.cellStyle = headerStyle
        sheet.setColumnWidth(idx, 6500)
    }

    report.items.forEachIndexed { index, item ->
        val row = sheet.createRow(index + 1)
        val rowStyle = when (item.verdict) {
            AuditVerdict.CRITICAL_NOT_CANCELLED, AuditVerdict.CRITICAL_NOT_ASSIGNED -> errorStyle
            AuditVerdict.ERROR_TIME_OR_STOP, AuditVerdict.NEEDS_MANUAL_CHECK -> warnStyle
            AuditVerdict.EXECUTED_CORRECTLY -> okStyle
            AuditVerdict.UNCHANGED_BASE_SCHEDULE -> null
        }

        val values = listOf(
            item.trainDisplay,
            item.verdict.badgeText,
            item.directiveTitle,
            item.telegramLocations.joinToString(", "),
            item.scheduleLocations.joinToString(", "),
            item.telegramTargetTimes.joinToString(", "),
            item.scheduleActualTimes.joinToString(", "),
            item.discrepancySummary,
            item.requiredAction
        )

        values.forEachIndexed { colIdx, text ->
            val cell = row.createCell(colIdx)
            cell.setCellValue(text)
            if (rowStyle != null) {
                cell.cellStyle = rowStyle
            }
        }
    }

    context.contentResolver.openOutputStream(targetUri)?.use { out ->
        workbook.write(out)
    }
    workbook.close()
}

fun getFileName(context: Context, uri: Uri): String {
    var result: String? = null
    if (uri.scheme == "content") {
        context.contentResolver.query(uri, null, null, null, null)?.use { cursor ->
            if (cursor.moveToFirst()) {
                val index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if (index >= 0) result = cursor.getString(index)
            }
        }
    }
    if (result == null) {
        result = uri.path
        val cut = result?.lastIndexOf('/') ?: -1
        if (cut != -1) result = result?.substring(cut + 1)
    }
    return result ?: "Выбранный файл"
}
