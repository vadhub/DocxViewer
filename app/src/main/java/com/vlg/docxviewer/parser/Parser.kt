package com.vlg.docxviewer.parser

import android.content.res.Resources
import android.util.Log
import android.util.Xml
import org.xmlpull.v1.XmlPullParser
import java.io.InputStream
import java.io.StringReader
import java.util.zip.ZipInputStream
import kotlin.math.roundToInt

class Parser(private val resources: Resources) {
    companion object {
        const val TAG = "DocxParser"

        private val wingdingsMapping = mapOf(
            "" to "•",   // U+F0B7 → U+2022
            "" to "◦",   // U+F0A8 → U+25E6
            "" to "▪",   // U+F09F → U+25AA
            "" to "–"    // U+F0D1 → U+2013
        )

        fun toRoman(num: Int): String {
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

        fun getListMarker(
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
    }

    fun parseDocx(inputStream: InputStream): ParseResult {
        val zip = ZipInputStream(inputStream)
        val elements = mutableListOf<DocxElement>()
        val images = mutableMapOf<String, ByteArray>()
        var xmlContent = ""
        var numberingXml = ""
        var relsXml = ""

        try {
            while (true) {
                val entry = zip.nextEntry ?: break
                when {
                    entry.name == "word/document.xml" ->
                        xmlContent = zip.readBytes().toString(Charsets.UTF_8)

                    entry.name == "word/numbering.xml" ->
                        numberingXml = zip.readBytes().toString(Charsets.UTF_8)

                    entry.name == "word/_rels/document.xml.rels" ->
                        relsXml = zip.readBytes().toString(Charsets.UTF_8)

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
        val imageRelationships = parseDocumentRels(relsXml)
        val listDefinitions = parseNumberingXml(numberingXml)
        val parseResult = ParseResult(elements, listDefinitions)
        parseDocumentXml(xmlContent, parseResult, images, imageRelationships)
        return parseResult
    }

    private fun parseDocumentRels(xml: String): Map<String, String> {
        val relationships = mutableMapOf<String, String>()
        if (xml.isEmpty()) return relationships

        val parser = Xml.newPullParser().apply {
            setFeature(XmlPullParser.FEATURE_PROCESS_NAMESPACES, true)
            setInput(StringReader(xml))
        }

        var eventType = parser.eventType
        while (eventType != XmlPullParser.END_DOCUMENT) {
            if (eventType == XmlPullParser.START_TAG && parser.name == "Relationship") {
                val id = parser.getAttributeValue(null, "Id")
                val target = parser.getAttributeValue(null, "Target")
                if (id != null && target != null) {
                    relationships[id] = target.substringAfter("media/")
                }
            }
            eventType = parser.next()
        }
        return relationships
    }

    private fun parseDocumentXml(
        xml: String,
        result: ParseResult,
        images: Map<String, ByteArray>,
        imageRelationships: Map<String, String>
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
        val drawingNS2 = "http://schemas.openxmlformats.org/drawingml/2006/main"
        val drawingNS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        val relationshipsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        var imageId = ""
        var imageRelId = ""
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
                                val colspan =
                                    parser.getAttributeValue(wordNS, "gridSpan")?.toIntOrNull() ?: 1
                                val rowspan = if (parser.getAttributeValue(
                                        wordNS,
                                        "vMerge"
                                    ) == "restart"
                                ) 1 else 0
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

                    drawingNS -> {
                        if (parser.name == "docPr") {
                            imageId = parser.getAttributeValue(null, "id") ?: ""
                        }
                    }

                    drawingNS2 -> {
                        if (parser.name == "blip") {
                            imageRelId = parser.getAttributeValue(relationshipsNS, "embed") ?: ""
                        }
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
                                    cellElements.add(
                                        DocxElement(
                                            ElementType.TEXT,
                                            text = text,
                                            textStyle = currentTextStyle
                                        )
                                    )
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

                        "r" -> {
                            if (imageId.isNotEmpty()) {
                                images["image$imageId.png"]?.let {
                                    result.elements.add(
                                        DocxElement(
                                            ElementType.IMAGE,
                                            imageData = it
                                        )
                                    )
                                }
                                imageId = ""
                            }

                            if (imageRelId.isNotEmpty()) {
                                val imageName = imageRelationships[imageRelId]
                                Log.d("!!!", imageRelId.toString())
                                Log.d("!!!", imageName.toString())
                                imageName?.let { name ->
                                    images[name]?.let { data ->
                                        result.elements.add(
                                            DocxElement(
                                                ElementType.IMAGE,
                                                imageData = data
                                            )
                                        )
                                    }
                                }
                                imageRelId = ""
                            }
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

                    "b" -> style = style.copy(isBold = true)
                    "i" -> style = style.copy(isItalic = true)
                    "u" -> style = style.copy(isUnderscore = true)
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


}