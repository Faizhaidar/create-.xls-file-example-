package com.example.faiz

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.widget.Button
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    var btn : Button?=null
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        btn = findViewById(R.id.bt_click)

        val jsonArray = JSONArray()
        val jsonObject = JSONObject()

        jsonObject.put("accountNumber", "")
        jsonObject.put("amount", "1000")
        jsonObject.put("date", "2023-02-07")
        jsonObject.put("expense_category", "Salaries")
        jsonObject.put("expense_type", "2")
        jsonObject.put("id", 1)
        jsonObject.put("notes", "Rent ")
        jsonObject.put("paidby", "Bank")
        jsonObject.put("time", "")
        jsonArray.put(jsonObject)

        val data = jsonArray.toString()

        btn?.setOnClickListener {
            saveToExcel(data)
        }
    }
    fun saveToExcel(jsonArray: String) {
        //val file = File(this.filesDir, "file_name.xlsx")

        val workbook = HSSFWorkbook()
        val sheet = workbook.createSheet("Expenses")

        val headers = arrayOf("Account Number", "Amount", "Date", "Expense Category", "Expense Type", "ID", "Notes", "Paid By", "Time")
        val headerRow = sheet.createRow(0)

        //set a header
        for (i in headers.indices) {
            headerRow.createCell(i).setCellValue(headers[i])
        }

        //set a data into cell in excel
        val jsonData = JSONArray(jsonArray)
        for (i in 0 until jsonData.length()) {
            val expense = jsonData.getJSONObject(i)
            val row = sheet.createRow(i + 1)
            row.createCell(0).setCellValue(expense.getString("accountNumber"))
            row.createCell(1).setCellValue(expense.getString("amount"))
            row.createCell(2).setCellValue(expense.getString("date"))
            row.createCell(3).setCellValue(expense.getString("expense_category"))
            row.createCell(4).setCellValue(expense.getString("expense_type"))
            row.createCell(5).setCellValue(expense.getString("id"))
            row.createCell(6).setCellValue(expense.getString("notes"))
            row.createCell(7).setCellValue(expense.getString("paidby"))
            row.createCell(8).setCellValue(expense.getString("time"))
        }

        val fileName = "expenses.xlsx"
        val file = File(getInternalStorage(), fileName)
        val stream = FileOutputStream(file)
        workbook.write(stream)
        stream.close()

        val intent = Intent(Intent.ACTION_SEND)
        intent.type = "application/vnd.ms-excel"
        intent.putExtra(Intent.EXTRA_STREAM, Uri.fromFile(file))
        startActivity(Intent.createChooser(intent, "Share Excel file"))
    }
    private fun getInternalStorage(): String {
        return this.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS).toString()
    }
}