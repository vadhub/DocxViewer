package com.vlg.docxviewer

import android.graphics.BitmapFactory
import android.graphics.Typeface
import android.os.Bundle
import android.util.Log
import android.util.TypedValue
import android.util.Xml
import android.view.Gravity
import android.view.ViewGroup
import android.widget.ImageView
import android.widget.LinearLayout
import android.widget.Space
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import org.xmlpull.v1.XmlPullParser
import java.io.InputStream
import java.io.StringReader
import java.util.zip.ZipInputStream
import kotlin.math.roundToInt

class MainActivity : AppCompatActivity() {
    companion object {
        private const val TAG = "DocxParser"
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val container = findViewById<LinearLayout>(R.id.container)
        val inputStream = resources.openRawResource(R.raw.test2)
        val result = parseDocx(inputStream)

        result.elements.forEach { element ->
            when (element.type) {
                ElementType.TEXT -> addTextView(container, element.text, element.textStyle)
                ElementType.IMAGE -> element.imageData?.let { addImageView(container, it) }
                ElementType.LIST_ITEM -> addListItem(
                    container,
                    element,
                    result.listDefinitions,
                    result.listCounters
                )

                ElementType.NEWLINE -> addSpace(container)
                ElementType.TABLE -> addTableView(container, element.table!!)
            }
        }
    }

    private val wingdingsMapping = mapOf(
        "" to "•",   // U+F0B7 → U+2022
        "" to "◦",   // U+F0A8 → U+25E6
        "" to "▪",   // U+F09F → U+25AA
        "" to "–"    // U+F0D1 → U+2013
    )

    data class DocxElement(
        val type: ElementType,
        val text: String? = null,
        val imageData: ByteArray? = null,
        val listLevel: Int = -1,
        val numId: Int = -1,
        val textStyle: TextStyle = TextStyle(),
        val table: DocxTable? = null
    )

    data class TextStyle(
        val fontSize: Float = 14f, // Размер в SP
        val isBold: Boolean = false,
        val isItalic: Boolean = false
    )

    data class ListStyle(val format: String, val lvlText: String)

    enum class ElementType { TEXT, IMAGE, LIST_ITEM, NEWLINE, TABLE }

    data class TableCell(
        val content: List<DocxElement>,
        val colspan: Int = 1,
        val rowspan: Int = 1
    )

    data class TableRowDocx(val cells: MutableList<TableCell>)
    data class DocxTable(val rows: MutableList<TableRowDocx>)

    data class ParseResult(
        val elements: MutableList<DocxElement>,
        val listDefinitions: Map<Int, Map<Int, ListStyle>> = emptyMap(),
        val listCounters: ListCounters = ListCounters()
    )

    class ListCounters {
        private val counters = mutableMapOf<String, Int>()

        fun getAndIncrement(numId: Int, level: Int): Int {
            val key = "${numId}_$level"
            val current = counters[key] ?: 1
            counters[key] = current + 1
            return current
        }

        fun resetUpperLevels(numId: Int, currentLevel: Int) {
            val keysToRemove = counters.keys.filter {
                val parts = it.split("_")
                parts[0].toInt() == numId && parts[1].toInt() > currentLevel
            }
            keysToRemove.forEach { counters.remove(it) }
        }
    }

    private fun parseDocx(inputStream: InputStream): ParseResult {
        val zip = ZipInputStream(inputStream)
        val elements = mutableListOf<DocxElement>()
        val images = mutableMapOf<String, ByteArray>()
        var xmlContent = ""
        var numberingXml = ""

        try {
            while (true) {
                val entry = zip.nextEntry ?: break
                when {
                    entry.name == "word/document.xml" ->
                        xmlContent = zip.readBytes().toString(Charsets.UTF_8)

                    entry.name == "word/numbering.xml" ->
                        numberingXml = zip.readBytes().toString(Charsets.UTF_8)

                    entry.name.startsWith("word/media/") -> {
                        val name = entry.name.substringAfterLast("/")
                        images[name] = zip.readBytes()
                    }
                }
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error parsing ZIP file", e)
        } finally {
            zip.close()
        }

        val listDefinitions = parseNumberingXml(numberingXml)
        val parseResult = ParseResult(elements, listDefinitions)
        parseDocumentXml(xmlContent, parseResult, images)
        return parseResult
    }

    private fun parseNumberingXml(xml: String): Map<Int, Map<Int, ListStyle>> {
        val numbering = mutableMapOf<Int, MutableMap<Int, ListStyle>>()
        if (xml.isEmpty()) return numbering

        val parser = Xml.newPullParser().apply {
            setFeature(XmlPullParser.FEATURE_PROCESS_NAMESPACES, true)
            setInput(StringReader(xml))
        }

        var eventType = parser.eventType
        var currentNumId = -1
        var currentAbstractNumId = -1
        var currentLevel = -1
        var currentFormat = "bullet"
        var currentLvlText = "•"
        val abstractNumMap = mutableMapOf<Int, MutableMap<Int, ListStyle>>()

        while (eventType != XmlPullParser.END_DOCUMENT) {
            when (eventType) {
                XmlPullParser.START_TAG -> when (parser.name) {
                    "num" -> currentNumId = parser.getAttributeValue(null, "numId")?.toInt() ?: -1
                    "abstractNum" -> currentAbstractNumId =
                        parser.getAttributeValue(null, "abstractNumId")?.toInt() ?: -1

                    "lvl" -> currentLevel = parser.getAttributeValue(null, "ilvl")?.toInt() ?: -1
                    "numFmt" -> currentFormat =
                        parser.getAttributeValue(null, "val") ?: currentFormat

                    "lvlText" -> {
                        currentLvlText = parser.getAttributeValue(null, "val")
                            ?.replace(Regex("%\\d+"), "")
                            ?.replace(Regex("&#x([0-9A-F]{4});")) { match ->
                                val hexCode = match.groupValues[1]
                                String(Character.toChars(hexCode.toInt(16)))
                            }
                            ?.let { text ->
                                wingdingsMapping[text] ?: text.takeIf { it.isNotBlank() } ?: "•"
                            }
                            ?: "•"
                    }

                    "abstractNumId" -> {
                        val abstractId = parser.getAttributeValue(null, "val")?.toInt() ?: -1
                        abstractNumMap[abstractId]?.let {
                            numbering[currentNumId] = it.toMutableMap()
                        }
                    }
                }

                XmlPullParser.END_TAG -> when (parser.name) {
                    "lvl" -> if (currentAbstractNumId != -1 && currentLevel != -1) {
                        abstractNumMap.getOrPut(currentAbstractNumId) { mutableMapOf() }[currentLevel] =
                            ListStyle(currentFormat, currentLvlText)
                        currentLevel = -1
                    }

                    "abstractNum" -> currentAbstractNumId = -1
                    "num" -> currentNumId = -1
                }
            }
            eventType = parser.next()
        }
        return numbering
    }

    private fun parseDocumentXml(
        xml: String,
        result: ParseResult,
        images: Map<String, ByteArray>
    ) {
        val parser = Xml.newPullParser().apply {
            setFeature(XmlPullParser.FEATURE_PROCESS_NAMESPACES, true)
            setInput(StringReader(xml))
        }

        var eventType = parser.eventType
        var currentParagraphText = StringBuilder()
        var currentListLevel = -1
        var currentNumId = -1
        var insideNumPr = false
        var currentTextStyle = TextStyle()
        val wordNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        val drawingNS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        var imageId = ""
        var previousNumId = -1
        var currentTable: DocxTable? = null
        var currentRow: TableRowDocx? = null
        var currentCell: TableCell? = null
        var cellElements = mutableListOf<DocxElement>()

        while (eventType != XmlPullParser.END_DOCUMENT) {
            when (eventType) {
                XmlPullParser.START_TAG -> when (parser.namespace) {
                    wordNS -> {
                        when (parser.name) {
                            "tbl" -> currentTable = DocxTable(mutableListOf())
                            "tr" -> currentRow = TableRowDocx(mutableListOf())
                            "tc" -> {
                                cellElements = mutableListOf()
                                // Парсим атрибуты ячейки
                                val colspan = parser.getAttributeValue(wordNS, "gridSpan")?.toIntOrNull() ?: 1
                                val rowspan = if (parser.getAttributeValue(wordNS, "vMerge") == "restart") 1 else 0
                                currentCell = TableCell(mutableListOf(), colspan, rowspan)
                            }

                            "rPr" -> currentTextStyle = parseRunProperties(parser)
                            "p" -> currentParagraphText.clear()
                            "t" -> currentParagraphText.append(parser.nextText())
                            "br" -> currentParagraphText.append("\n")
                            "tab" -> currentParagraphText.append("\t")
                            "numPr" -> insideNumPr = true
                            "ilvl" -> if (insideNumPr) {
                                currentListLevel =
                                    parser.getAttributeValue(null, "val")?.toInt() ?: 0
                            }

                            "numId" -> if (insideNumPr) {
                                currentNumId = parser.getAttributeValue(null, "val")?.toInt() ?: -1
                            }
                        }
                    }

                    drawingNS -> if (parser.name == "docPr") {
                        imageId = parser.getAttributeValue(null, "id") ?: ""
                    }

                }

                XmlPullParser.END_TAG -> when (parser.namespace) {
                    wordNS -> when (parser.name) {
                        "tbl" -> {
                            currentTable?.let {
                                result.elements.add(DocxElement(ElementType.TABLE, table = it))
                                currentTable = null
                            }
                        }
                        "tr" -> {
                            currentRow?.let {
                                currentTable?.rows?.add(it)
                                currentRow = null
                            }
                        }
                        "tc" -> {
                            currentCell?.let {
                                currentRow?.cells?.add(it.copy(content = cellElements))
                                currentCell = null
                                cellElements = mutableListOf()
                            }
                        }

                        "p" -> {
                            if (currentCell != null) {
                                val text = currentParagraphText.toString().trim()
                                if (text.isNotEmpty()) {
                                    cellElements.add(DocxElement(
                                        ElementType.TEXT,
                                        text = text,
                                        textStyle = currentTextStyle
                                    ))
                                }
                            } else if (currentListLevel != -1 && currentNumId != -1) {
                                if (previousNumId != currentNumId) {
                                    result.listCounters.resetUpperLevels(
                                        previousNumId,
                                        currentListLevel
                                    )
                                    previousNumId = currentNumId
                                }
                                result.elements.add(
                                    DocxElement(
                                        ElementType.LIST_ITEM,
                                        text = currentParagraphText.toString().trim(),
                                        listLevel = currentListLevel,
                                        numId = currentNumId,
                                        textStyle = currentTextStyle

                                    )
                                )
                            } else if (currentParagraphText.isNotEmpty()) {
                                result.elements.add(
                                    DocxElement(
                                        ElementType.TEXT,
                                        currentParagraphText.toString().trim(),
                                        textStyle = currentTextStyle
                                    )
                                )
                                result.elements.add(DocxElement(ElementType.NEWLINE))
                            }
                            currentParagraphText.clear()
                            currentListLevel = -1
                            currentNumId = -1
                        }

                        "r" -> if (imageId.isNotEmpty()) {
                            images["image$imageId.png"]?.let {
                                result.elements.add(DocxElement(ElementType.IMAGE, imageData = it))
                            }
                            imageId = ""
                        }

                        "numPr" -> insideNumPr = false
                    }
                }
            }
            eventType = parser.next()
        }
    }

    private fun parseRunProperties(parser: XmlPullParser): TextStyle {
        var style = TextStyle()
        var eventType = parser.next()

        while (!(eventType == XmlPullParser.END_TAG && parser.name == "rPr")) {
            if (eventType == XmlPullParser.START_TAG) {
                when (parser.name) {
                    "sz" -> {
                        val halfPoints =
                            parser.getAttributeValue(null, "val")?.toFloatOrNull() ?: 24f
                        style = style.copy(fontSize = halfPointsToSp(halfPoints))
                    }

                    "b" -> {
                        style = style.copy(isBold = true)
                    }

                    "i" -> style = style.copy(isItalic = true)
                }
            }
            eventType = parser.next()
        }
        return style
    }

    private fun halfPointsToSp(halfPoints: Float): Float {
        val points = halfPoints / 2
        return (points * resources.displayMetrics.scaledDensity).roundToInt().toFloat()
    }

    private fun TextView.applyTextStyle(style: TextStyle) {
        setTextSize(TypedValue.COMPLEX_UNIT_SP, style.fontSize)
        val typeface = when {
            style.isBold && style.isItalic -> Typeface.BOLD_ITALIC
            style.isBold -> Typeface.BOLD
            style.isItalic -> Typeface.ITALIC
            else -> Typeface.NORMAL
        }
        setTypeface(null, typeface)

    }

    private fun addTextView(container: LinearLayout, text: String?, element: TextStyle) {
        if (text.isNullOrEmpty()) return

        TextView(this).apply {
            this.text = text.replace("\n", "")
//            setTextColor(Color.BLACK)
            applyTextStyle(element)
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
            ).apply {
                setTextIsSelectable(true)
                setMargins(16.dp(), 4.dp(), 16.dp(), 4.dp())
            }
            container.addView(this)
        }
    }

    private fun addTableView(container: LinearLayout, table: DocxTable) {
        val tableLayout = TableLayout(this).apply {
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
            ).apply {
                setMargins(16.dp(), 8.dp(), 16.dp(), 8.dp())
            }
            setBackgroundResource(R.drawable.cell_border)
        }

        // Рассчитываем ширину столбцов
        val columnWeights = calculateColumnWeights(table)

        table.rows.forEach { row ->
            val tableRow = TableRow(this).apply {
                layoutParams = TableLayout.LayoutParams(
                    TableLayout.LayoutParams.MATCH_PARENT,
                    TableLayout.LayoutParams.WRAP_CONTENT
                )
            }

            row.cells.forEachIndexed { index, cell ->
                val cellLayout = LinearLayout(this).apply {
                    orientation = LinearLayout.VERTICAL
                    layoutParams = TableRow.LayoutParams(
                        0,
                        TableRow.LayoutParams.WRAP_CONTENT,
                        columnWeights.getOrElse(index) { 1f }
                    ).apply {
                        marginEnd = 2.dp()
                        bottomMargin = 2.dp()
                        gravity = Gravity.CENTER_VERTICAL
                    }
                    setPadding(8.dp(), 8.dp(), 8.dp(), 8.dp())
                    setBackgroundResource(R.drawable.cell_border)
                }

                cell.content.forEach { element ->
                    when (element.type) {
                        ElementType.TEXT -> addTextView(cellLayout, element.text, element.textStyle)
                        ElementType.IMAGE -> element.imageData?.let { addImageView(cellLayout, it) }
                        ElementType.LIST_ITEM -> addListItem(
                            cellLayout,
                            element
                        )
                        ElementType.NEWLINE -> addSpace(cellLayout)
                        else -> {}
                    }
                }

                // Обработка объединенных ячеек
                if (cell.colspan > 1 || cell.rowspan > 1) {
                    val params = cellLayout.layoutParams as TableRow.LayoutParams
                    params.span = cell.colspan
                }

                tableRow.addView(cellLayout)
            }
            tableLayout.addView(tableRow)
        }
        container.addView(tableLayout)
    }

    private fun calculateColumnWeights(table: DocxTable): List<Float> {
        val maxColumns = table.rows.maxOfOrNull { it.cells.size } ?: 0
        val columnWeights = MutableList(maxColumns) { 0f }

        table.rows.forEach { row ->
            row.cells.forEachIndexed { index, cell ->
                val weight = cell.colspan.toFloat() / row.cells.sumOf { it.colspan.toDouble() }
                columnWeights[index] = columnWeights[index].coerceAtLeast(weight.toFloat())
            }
        }

        return columnWeights
    }


    private fun addImageView(container: LinearLayout, data: ByteArray) {
        ImageView(this).apply {
            setImageBitmap(BitmapFactory.decodeByteArray(data, 0, data.size))
            adjustViewBounds = true
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
            ).apply {
                setMargins(16.dp(), 8.dp(), 16.dp(), 8.dp())
            }
            container.addView(this)
        }
    }

    private fun addListItem(
        container: LinearLayout,
        element: DocxElement,
        listDefinitions: Map<Int, Map<Int, ListStyle>> = mutableMapOf(),
        counters: ListCounters = ListCounters()
    ) {
        val markerText = if (listDefinitions.isNotEmpty()) getListMarker(element, listDefinitions, counters) else getDefaultListMarker()

        LinearLayout(this).apply {
            orientation = LinearLayout.HORIZONTAL
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
            ).apply {
                setMargins(
                    (16 + element.listLevel * 24).dp(),
                    4.dp(),
                    16.dp(),
                    4.dp()
                )
            }

            TextView(context).apply {
                text = markerText
//                setTextColor(Color.BLACK)
                setTextIsSelectable(true)
                typeface = Typeface.DEFAULT
                textSize = 16f
                layoutParams = LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.WRAP_CONTENT,
                    ViewGroup.LayoutParams.WRAP_CONTENT
                ).apply {
                    marginEnd = 8.dp()
                }
                addView(this)
            }

            TextView(context).apply {
                setTextIsSelectable(true)
                text = element.text
//                setTextColor(Color.BLACK)
                layoutParams = LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.MATCH_PARENT,
                    ViewGroup.LayoutParams.WRAP_CONTENT
                )
                addView(this)
            }

            container.addView(this)
        }
    }

    private fun getDefaultListMarker() = "•"

    private fun getListMarker(
        element: DocxElement,
        listDefinitions: Map<Int, Map<Int, ListStyle>>,
        counters: ListCounters
    ): String {
        val levelStyles = listDefinitions[element.numId] ?: return "•"
        val style = levelStyles[element.listLevel] ?: return "•"

        return when (style.format) {
            "bullet" -> style.lvlText.takeIf { it.isNotEmpty() }?.let {
                wingdingsMapping[it] ?: it.first().toString()
            } ?: "•"

            "decimal" -> "${counters.getAndIncrement(element.numId, element.listLevel)}."
            "lowerLetter" -> "${
                ('a' + counters.getAndIncrement(
                    element.numId,
                    element.listLevel
                ) - 1)
            }."

            "upperLetter" -> "${
                ('A' + counters.getAndIncrement(
                    element.numId,
                    element.listLevel
                ) - 1)
            }."

            "lowerRoman" -> "${
                toRoman(
                    counters.getAndIncrement(
                        element.numId,
                        element.listLevel
                    )
                ).lowercase()
            }."

            "upperRoman" -> "${
                toRoman(
                    counters.getAndIncrement(
                        element.numId,
                        element.listLevel
                    )
                )
            }."

            else -> "•"
        }.also {
            counters.resetUpperLevels(element.numId, element.listLevel)
        }
    }

    private fun safeFormat(pattern: String, value: Int): String {
        return try {
            val safePattern = pattern.replace(Regex("%(?![ds])"), "%%")
            String.format(safePattern, value)
        } catch (e: Exception) {
            Log.e(TAG, "Format pattern error: $pattern", e)
            "$value."
        }
    }

    private fun toRoman(num: Int): String {
        if (num < 1 || num > 3999) return ""
        val values = intArrayOf(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
        val romanNumerals =
            arrayOf("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
        val sb = StringBuilder()
        var number = num
        for (i in values.indices) {
            while (number >= values[i]) {
                number -= values[i]
                sb.append(romanNumerals[i])
            }
        }
        return sb.toString()
    }

    private fun addSpace(container: LinearLayout) {
        container.addView(Space(this).apply {
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                8.dp()
            )
        })
    }

    private fun Int.dp() = (this * resources.displayMetrics.density).toInt()
}