package com.vlg.docxviewer

import androidx.fragment.app.Fragment

interface Navigator {
    fun startFragment(fragment: Fragment)
    fun finishActivity()
}