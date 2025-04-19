package com.vlg.docxviewer.parser

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