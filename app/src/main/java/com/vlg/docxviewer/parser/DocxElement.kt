package com.vlg.docxviewer.parser

data class DocxElement(
    val type: ElementType,
    val text: String? = null,
    val imageData: ByteArray? = null,
    val listLevel: Int = -1,
    val numId: Int = -1,
    val textStyle: TextStyle = TextStyle(),
    val table: DocxTable? = null
)
