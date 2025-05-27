package com.vlg.docxviewer

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.LinearLayout
import androidx.fragment.app.Fragment
import com.vlg.docxviewer.displayer.Displayer
import com.vlg.docxviewer.parser.Parser
import java.io.FileInputStream
import java.io.InputStream
import androidx.core.net.toUri
import java.io.FileNotFoundException

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

        try {
            val inputStream = context?.contentResolver?.openInputStream(Uri.parse(uri))
            inputStream?.use { stream ->
                val result = Parser(resources).parseDocx(stream)
                Displayer(view.context).display(result, container)
            }
        } catch (e: FileNotFoundException) {
            Log.e("FileError", "File not found: ${e.message}")
        } catch (e: SecurityException) {
            Log.e("FileError", "Permission denied: ${e.message}")
        }
    }
}