package com.vlg.docxviewer.parser

data class ParseResult(
    val elements: MutableList<DocxElement>,
    val listDefinitions: Map<Int, Map<Int, ListStyle>> = emptyMap(),
    val listCounters: ListCounters = ListCounters()
)