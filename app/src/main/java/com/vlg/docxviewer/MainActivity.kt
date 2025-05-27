package com.vlg.docxviewer

import android.content.Intent
import android.os.Bundle
import android.widget.LinearLayout
import androidx.appcompat.app.AppCompatActivity
import androidx.fragment.app.Fragment
import com.vlg.docxviewer.displayer.Displayer
import com.vlg.docxviewer.parser.Parser

class MainActivity : AppCompatActivity(), Navigator {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        val intent = intent
        if (Intent.ACTION_VIEW == intent.action) {
            val args = Bundle()
            args.putString("uri", intent.data.toString())
            val fragment = DocxViewerFragment()
            fragment.arguments = args
            startFragment(fragment)
        } else {
            savedInstanceState ?: startFragment(DocumentsFragment())
        }
    }

    override fun startFragment(fragment: Fragment) {
        supportFragmentManager.beginTransaction()
            .addToBackStack(null)
            .replace(R.id.container, fragment, "fragment")
            .commit()
    }

    override fun finishActivity() {
        finish()
    }


}