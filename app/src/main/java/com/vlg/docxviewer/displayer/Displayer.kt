package com.vlg.docxviewer.displayer

import android.content.Context
import android.graphics.BitmapFactory
import android.graphics.Paint
import android.graphics.Typeface
import android.util.TypedValue
import android.view.Gravity
import android.view.ViewGroup
import android.widget.ImageView
import android.widget.LinearLayout
import android.widget.Space
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import com.vlg.docxviewer.R
import com.vlg.docxviewer.parser.DocxElement
import com.vlg.docxviewer.parser.DocxTable
import com.vlg.docxviewer.parser.ElementType
import com.vlg.docxviewer.parser.ListCounters
import com.vlg.docxviewer.parser.ListStyle
import com.vlg.docxviewer.parser.ParseResult
import com.vlg.docxviewer.parser.Parser.Companion.getListMarker
import com.vlg.docxviewer.parser.TextStyle

class Displayer(private val context: Context) {

    fun display(result: ParseResult, container: LinearLayout) {
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

    private fun TextView.applyTextStyle(style: TextStyle) {
        setTextSize(TypedValue.COMPLEX_UNIT_SP, style.fontSize)
        val typeface = when {
            style.isBold && style.isItalic -> Typeface.BOLD_ITALIC
            style.isBold -> Typeface.BOLD
            style.isItalic -> Typeface.ITALIC
            else -> Typeface.NORMAL
        }
        if (style.isUnderscore) {
            paintFlags = Paint.UNDERLINE_TEXT_FLAG
        }
        setTypeface(null, typeface)

    }

    private fun addTextView(container: LinearLayout, text: String?, element: TextStyle) {
        if (text.isNullOrEmpty()) return

        TextView(context).apply {
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
        val tableLayout = TableLayout(context).apply {
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.WRAP_CONTENT
            ).apply {
                setMargins(16.dp(), 8.dp(), 16.dp(), 8.dp())
            }
            setBackgroundResource(R.drawable.cell_border)
        }

        val columnWeights = calculateColumnWeights(table)

        table.rows.forEachIndexed { rowIndex, row ->
            val tableRow = TableRow(context).apply {
                layoutParams = TableLayout.LayoutParams(
                    TableLayout.LayoutParams.MATCH_PARENT,
                    TableLayout.LayoutParams.WRAP_CONTENT
                )
            }

            row.cells.forEachIndexed { index, cell ->
                val cellLayout = LinearLayout(context).apply {
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
        ImageView(context).apply {
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

        LinearLayout(context).apply {
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

    private fun addSpace(container: LinearLayout) {
        container.addView(Space(context).apply {
            layoutParams = LinearLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                8.dp()
            )
        })
    }

    private fun Int.dp() = (this * context.resources.displayMetrics.density).toInt()
}