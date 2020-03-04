package com.xlsx

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.IndexedColors
import java.io.FileOutputStream


class MainActivity : AppCompatActivity() {

    companion object {
        val COLUMNS = arrayOf("First Name", "Last Name", "Age")
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
    }

    override fun onResume() {
        super.onResume()

        val persons = listOf<Person>(
            Person("Vinicius","Almada", 22),
            Person("Mohara","Nascimento", 23),
            Person("Pepita","Nascimento Almada", 1)
        )

        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Persons")
        val headerFont = workbook.createFont()
        headerFont.bold = true
        headerFont.fontHeightInPoints = 14.toShort()
        headerFont.color = IndexedColors.RED.getIndex()

        val headerCellStyle = workbook.createCellStyle()
        headerCellStyle.setFont(headerFont)

        // Create a Row
        val headerRow = sheet.createRow(0)

        for (i in COLUMNS.indices) {
            val cell = headerRow.createCell(i)
            cell.setCellValue(COLUMNS[i])
            cell.cellStyle = headerCellStyle
        }

        var rowNum = 1;

        persons.forEach {
            val r = sheet.createRow(rowNum++)
            r.createCell(0).setCellValue(it.firstName)
            r.createCell(1).setCellValue(it.lastName)
            r.createCell(2).setCellValue(it.age.toDouble())
        }

        COLUMNS.indices.forEach {
            sheet.autoSizeColumn(it)
        }

        val output = FileOutputStream("persons.xlxs")
        workbook.write(output)
        output.close()
    }
}
