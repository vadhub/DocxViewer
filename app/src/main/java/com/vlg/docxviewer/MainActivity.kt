package com.vlg.docxviewer

import android.os.Bundle
import android.widget.LinearLayout
import androidx.appcompat.app.AppCompatActivity
import com.vlg.docxviewer.displayer.Displayer
import com.vlg.docxviewer.parser.Parser

class MainActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val container = findViewById<LinearLayout>(R.id.container)
        val inputStream = resources.openRawResource(R.raw.test2)
        val result = Parser(resources).parseDocx(inputStream)
        Displayer(this).display(result, container)
    }


}