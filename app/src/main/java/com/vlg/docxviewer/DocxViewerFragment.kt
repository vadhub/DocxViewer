package com.vlg.docxviewer

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.LinearLayout
import androidx.fragment.app.Fragment
import com.vlg.docxviewer.displayer.Displayer
import com.vlg.docxviewer.parser.Parser

class DocxViewerFragment : Fragment() {

    override fun onCreateView(
        inflater: LayoutInflater,
        container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        return inflater.inflate(R.layout.fragment_docx_viewer, container, false)
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        val uri = arguments?.getString("uri")

        val container = view.findViewById<LinearLayout>(R.id.container)
        val inputStream = resources.openRawResource(R.raw.test2)
        val result = Parser(resources).parseDocx(inputStream)
        Displayer(view.context).display(result, container)
    }
}