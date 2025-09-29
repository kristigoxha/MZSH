package com.example.mzsh

import android.Manifest
import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.widget.*
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFTable
import java.io.*
import java.text.SimpleDateFormat
import java.util.*

data class Employee(
    val id: Int,
    val name: String,
    val position: String,
    val grade: String,
    val years: Int,
    val hireDate: String,
    var workDates: String,
    val workDays: Int,
    val overtime: String
)

class MainActivity : AppCompatActivity() {

    private lateinit var monthSpinner: Spinner
    private lateinit var yearEditText: EditText
    private lateinit var employeeRecyclerView: RecyclerView
    private lateinit var loadDocButton: Button
    private lateinit var saveDocButton: Button
    private lateinit var employeeAdapter: EmployeeAdapter

    private var employees = mutableListOf<Employee>()
    private var currentDocumentUri: Uri? = null
    private var originalDocument: XWPFDocument? = null

    private val documentPickerLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == Activity.RESULT_OK) {
            result.data?.data?.let { uri ->
                currentDocumentUri = uri
                loadDocument(uri)
            }
        }
    }

    private val documentSaveLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == Activity.RESULT_OK) {
            result.data?.data?.let { uri ->
                saveDocument(uri)
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        initViews()
        setupSpinner()
        setupRecyclerView()
        requestPermissions()
    }

    private fun initViews() {
        monthSpinner = findViewById(R.id.monthSpinner)
        yearEditText = findViewById(R.id.yearEditText)
        employeeRecyclerView = findViewById(R.id.employeeRecyclerView)
        loadDocButton = findViewById(R.id.loadDocButton)
        saveDocButton = findViewById(R.id.saveDocButton)

        loadDocButton.setOnClickListener { openDocumentPicker() }
        saveDocButton.setOnClickListener { saveDocumentAs() }

        // Set current year as default
        yearEditText.setText(Calendar.getInstance().get(Calendar.YEAR).toString())
    }

    private fun setupSpinner() {
        val months = arrayOf(
            "Janar", "Shkurt", "Mars", "Prill", "Maj", "Qershor",
            "Korrik", "Gusht", "Shtator", "Tetor", "Nëntor", "Dhjetor"
        )

        val adapter = ArrayAdapter(this, android.R.layout.simple_spinner_item, months)
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item)
        monthSpinner.adapter = adapter

        // Set current month as default
        monthSpinner.setSelection(Calendar.getInstance().get(Calendar.MONTH))
    }

    private fun setupRecyclerView() {
        employeeAdapter = EmployeeAdapter(employees) { employee, newDates ->
            updateEmployeeDates(employee, newDates)
        }
        employeeRecyclerView.layoutManager = LinearLayoutManager(this)
        employeeRecyclerView.adapter = employeeAdapter
    }

    private fun requestPermissions() {
        val permissions = arrayOf(
            Manifest.permission.READ_EXTERNAL_STORAGE,
            Manifest.permission.WRITE_EXTERNAL_STORAGE
        )

        val permissionsToRequest = permissions.filter {
            ContextCompat.checkSelfPermission(this, it) != PackageManager.PERMISSION_GRANTED
        }

        if (permissionsToRequest.isNotEmpty()) {
            ActivityCompat.requestPermissions(this, permissionsToRequest.toTypedArray(), 100)
        }
    }

    private fun openDocumentPicker() {
        val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
        documentPickerLauncher.launch(intent)
    }

    private fun loadDocument(uri: Uri) {
        try {
            contentResolver.openInputStream(uri)?.use { inputStream ->
                originalDocument = XWPFDocument(inputStream)
                parseDocument(originalDocument!!)
                Toast.makeText(this, "Dokumenti u ngarkua me sukses", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            Toast.makeText(this, "Gabim në ngarkimin e dokumentit: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun parseDocument(document: XWPFDocument) {
        employees.clear()

        // Find the table in the document
        val tables = document.tables
        if (tables.isNotEmpty()) {
            val table = tables[0] // Assuming the first table is our target

            // Skip header row (index 0) and start from row 1
            for (i in 1 until table.rows.size) {
                val row = table.getRow(i)
                if (row.tableCells.size >= 9) {
                    try {
                        val employee = Employee(
                            id = row.getCell(0).text.toIntOrNull() ?: i,
                            name = row.getCell(1).text,
                            position = row.getCell(2).text,
                            grade = row.getCell(3).text,
                            years = row.getCell(4).text.toIntOrNull() ?: 0,
                            hireDate = row.getCell(5).text,
                            workDates = row.getCell(6).text,
                            workDays = row.getCell(7).text.toIntOrNull() ?: 0,
                            overtime = row.getCell(8).text
                        )
                        employees.add(employee)
                    } catch (e: Exception) {
                        // Skip malformed rows
                        continue
                    }
                }
            }
        }

        employeeAdapter.notifyDataSetChanged()
    }

    private fun updateEmployeeDates(employee: Employee, newDates: String) {
        val index = employees.indexOf(employee)
        if (index != -1) {
            employees[index] = employee.copy(workDates = newDates)
        }
    }

    private fun saveDocumentAs() {
        if (originalDocument == null) {
            Toast.makeText(this, "Nuk ka dokument të ngarkuar", Toast.LENGTH_SHORT).show()
            return
        }

        val intent = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            val month = monthSpinner.selectedItem.toString()
            val year = yearEditText.text.toString()
            putExtra(Intent.EXTRA_TITLE, "Kohëshënuesi $month $year.docx")
        }
        documentSaveLauncher.launch(intent)
    }

    private fun saveDocument(uri: Uri) {
        try {
            val document = XWPFDocument(originalDocument!!.packagePart.inputStream)

            // Update document title
            val month = monthSpinner.selectedItem.toString()
            val year = yearEditText.text.toString()

            // Update the header text
            updateDocumentHeader(document, month, year)

            // Update table data
            updateTableData(document)

            contentResolver.openOutputStream(uri)?.use { outputStream ->
                document.write(outputStream)
                document.close()
                Toast.makeText(this, "Dokumenti u ruajt me sukses", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            Toast.makeText(this, "Gabim në ruajtjen e dokumentit: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun updateDocumentHeader(document: XWPFDocument, month: String, year: String) {
        // Update paragraphs that contain the month/year info
        for (paragraph in document.paragraphs) {
            val text = paragraph.text
            if (text.contains("muajin")) {
                // Replace the month in the title
                val runs = paragraph.runs
                for (run in runs) {
                    val runText = run.text()
                    if (runText != null && runText.contains("muajin")) {
                        run.setText("Kohëshënuesi i MZSH-së Memaliaj për muajin $month $year", 0)
                    }
                }
            }
        }
    }

    private fun updateTableData(document: XWPFDocument) {
        val tables = document.tables
        if (tables.isNotEmpty()) {
            val table = tables[0]

            // Update table rows with new data
            for (i in 1 until table.rows.size) {
                if (i - 1 < employees.size) {
                    val employee = employees[i - 1]
                    val row = table.getRow(i)

                    // Update the work dates column (index 6)
                    if (row.tableCells.size > 6) {
                        row.getCell(6).text = employee.workDates
                    }
                }
            }
        }
    }
}

// EmployeeAdapter.kt
class EmployeeAdapter(
    private val employees: List<Employee>,
    private val onDatesChanged: (Employee, String) -> Unit
) : RecyclerView.Adapter<EmployeeAdapter.EmployeeViewHolder>() {

    class EmployeeViewHolder(itemView: android.view.View) : RecyclerView.ViewHolder(itemView) {
        val nameTextView: TextView = itemView.findViewById(R.id.nameTextView)
        val positionTextView: TextView = itemView.findViewById(R.id.positionTextView)
        val datesEditText: EditText = itemView.findViewById(R.id.datesEditText)
        val workDaysTextView: TextView = itemView.findViewById(R.id.workDaysTextView)
    }

    override fun onCreateViewHolder(parent: android.view.ViewGroup, viewType: Int): EmployeeViewHolder {
        val view = android.view.LayoutInflater.from(parent.context)
            .inflate(R.layout.item_employee, parent, false)
        return EmployeeViewHolder(view)
    }

    override fun onBindViewHolder(holder: EmployeeViewHolder, position: Int) {
        val employee = employees[position]

        holder.nameTextView.text = employee.name
        holder.positionTextView.text = employee.position
        holder.datesEditText.setText(employee.workDates)
        holder.workDaysTextView.text = "${employee.workDays} ditë"

        holder.datesEditText.setOnFocusChangeListener { _, hasFocus ->
            if (!hasFocus) {
                val newDates = holder.datesEditText.text.toString()
                onDatesChanged(employee, newDates)
            }
        }
    }

    override fun getItemCount() = employees.size
}