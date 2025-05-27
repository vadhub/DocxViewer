package com.vlg.docxviewer

import android.content.Context
import android.content.Intent
import android.os.Bundle
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import androidx.appcompat.app.AppCompatActivity
import androidx.fragment.app.Fragment

class DocumentsFragment : Fragment() {

    private val REQUEST_CODE_OPEN_DIRECTORY = 134

    private lateinit var navigator: Navigator

    override fun onAttach(context: Context) {
        super.onAttach(context)
        navigator = context as MainActivity
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View {
        return inflater.inflate(R.layout.fragment_document,  container, false)
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        val intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.setType("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        startActivityForResult(intent, REQUEST_CODE_OPEN_DIRECTORY)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        if (requestCode == REQUEST_CODE_OPEN_DIRECTORY && resultCode == AppCompatActivity.RESULT_OK) {
            Log.d("ddd", String.format("Open Directory result Uri : %s", data!!.data))
            val args = Bundle()
            args.putString("uri", data.data.toString())
            val fragment = DocxViewerFragment()
            fragment.arguments = args
            navigator.startFragment(fragment)
        } else {
            navigator.finishActivity()
        }
    }

}