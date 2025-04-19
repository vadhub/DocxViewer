package com.vlg.docxviewer.parser

data class TableCell(
    val content: List<DocxElement>,
    val colspan: Int = 1,
    val rowspan: Int = 1
)

data class TableRowDocx(val cells: MutableList<TableCell>)
data class DocxTable(val rows: MutableList<TableRowDocx>)
