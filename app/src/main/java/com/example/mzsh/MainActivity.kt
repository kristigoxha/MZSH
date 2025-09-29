package com.example.mzsh

import android.Manifest
import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.widget.*
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.*
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
        // Android 11+ requires MANAGE_EXTERNAL_STORAGE for broad access
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            // Using Storage Access Framework (SAF) - no special permissions needed
            return
        }
        
        // For Android 10 and below
        val permissions = mutableListOf<String>()
        if (Build.VERSION.SDK_INT <= Build.VERSION_CODES.P) {
            permissions.add(Manifest.permission.READ_EXTERNAL_STORAGE)
            permissions.add(Manifest.permission.WRITE_EXTERNAL_STORAGE)
        }

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
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }
        documentPickerLauncher.launch(intent)
    }

    private fun loadDocument(uri: Uri) {
        try {
            contentResolver.openInputStream(uri)?.use { inputStream ->
                // Create a copy to avoid stream closure issues
                val byteArray = inputStream.readBytes()
                originalDocument = XWPFDocument(ByteArrayInputStream(byteArray))
                parseDocument(originalDocument!!)
                Toast.makeText(this, "Dokumenti u ngarkua me sukses", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Gabim në ngarkimin e dokumentit: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun parseDocument(document: XWPFDocument) {
        employees.clear()

        try {
            // Find the table in the document
            val tables = document.tables
            if (tables.isEmpty()) {
                Toast.makeText(this, "Nuk u gjet tabela në dokument", Toast.LENGTH_SHORT).show()
                return
            }

            val table = tables[0] // Assuming the first table is our target

            // Skip header row (index 0) and start from row 1
            for (i in 1 until table.rows.size) {
                val row = table.getRow(i)
                if (row.tableCells.size >= 9) {
                    try {
                        val employee = Employee(
                            id = row.getCell(0).text.trim().toIntOrNull() ?: i,
                            name = row.getCell(1).text.trim(),
                            position = row.getCell(2).text.trim(),
                            grade = row.getCell(3).text.trim(),
                            years = row.getCell(4).text.trim().toIntOrNull() ?: 0,
                            hireDate = row.getCell(5).text.trim(),
                            workDates = row.getCell(6).text.trim(),
                            workDays = row.getCell(7).text.trim().toIntOrNull() ?: 0,
                            overtime = row.getCell(8).text.trim()
                        )
                        employees.add(employee)
                    } catch (e: Exception) {
                        // Skip malformed rows
                        e.printStackTrace()
                        continue
                    }
                }
            }

            employeeAdapter.notifyDataSetChanged()
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Gabim në leximin e tabelës: ${e.message}", Toast.LENGTH_LONG).show()
        }
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
            putExtra(Intent.EXTRA_TITLE, "Kohëshënuesi_${month}_${year}.docx")
            addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
        }
        documentSaveLauncher.launch(intent)
    }

    private fun saveDocument(uri: Uri) {
        try {
            if (originalDocument == null) {
                Toast.makeText(this, "Dokumenti origjinal mungon", Toast.LENGTH_SHORT).show()
                return
            }

            // Create a new document from the original
            val tempBytes = ByteArrayOutputStream()
            originalDocument!!.write(tempBytes)
            val document = XWPFDocument(ByteArrayInputStream(tempBytes.toByteArray()))

            // Update document title
            val month = monthSpinner.selectedItem.toString()
            val year = yearEditText.text.toString()

            // Update the header text
            updateDocumentHeader(document, month, year)

            // Update table data
            updateTableData(document)

            // Save to file
            contentResolver.openOutputStream(uri)?.use { outputStream ->
                document.write(outputStream)
                document.close()
                Toast.makeText(this, "Dokumenti u ruajt me sukses", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Gabim në ruajtjen e dokumentit: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun updateDocumentHeader(document: XWPFDocument, month: String, year: String) {
        try {
            // Update paragraphs that contain the month/year info
            for (paragraph in document.paragraphs) {
                val text = paragraph.text
                if (text.contains("muajin", ignoreCase = true)) {
                    // Clear existing runs and create new one
                    while (paragraph.runs.isNotEmpty()) {
                        paragraph.removeRun(0)
                    }
                    val run = paragraph.createRun()
                    run.setText("Kohëshënuesi i MZSH-së Memaliaj për muajin $month $year")
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    private fun updateTableData(document: XWPFDocument) {
        try {
            val tables = document.tables
            if (tables.isEmpty()) return

            val table = tables[0]

            // Update table rows with new data
            for (i in 1 until table.rows.size) {
                if (i - 1 < employees.size) {
                    val employee = employees[i - 1]
                    val row = table.getRow(i)

                    // Update the work dates column (index 6)
                    if (row.tableCells.size > 6) {
                        val cell = row.getCell(6)
                        // Clear existing content
                        while (cell.paragraphs.size > 1) {
                            cell.removeParagraph(1)
                        }
                        if (cell.paragraphs.isNotEmpty()) {
                            val paragraph = cell.paragraphs[0]
                            while (paragraph.runs.isNotEmpty()) {
                                paragraph.removeRun(0)
                            }
                            val run = paragraph.createRun()
                            run.setText(employee.workDates)
                        }
                    }
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    override fun onDestroy() {
        super.onDestroy()
        try {
            originalDocument?.close()
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}

// EmployeeAdapter
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

        // Update on focus loss
        holder.datesEditText.setOnFocusChangeListener { _, hasFocus ->
            if (!hasFocus) {
                val newDates = holder.datesEditText.text.toString()
                if (newDates != employee.workDates) {
                    onDatesChanged(employee, newDates)
                }
            }
        }
    }

    override fun getItemCount() = employees.size
}