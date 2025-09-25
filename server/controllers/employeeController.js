const Employee = require('../models/Employee');
const XLSX = require('xlsx'); // for Excel export
const fs = require('fs');
const path = require('path');

// ------------------- CRUD OPERATIONS -------------------

// GET all employees
exports.getAll = async (req, res) => {
  try {
    const employees = await Employee.find().lean();
    res.json(employees);
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: err.message });
  }
};

// GET single employee by ID
exports.getOne = async (req, res) => {
  try {
    const emp = await Employee.findById(req.params.id).lean();
    if (!emp) return res.status(404).json({ message: 'Employee not found' });

    res.json(emp);
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Server error while fetching employee' });
  }
};

// CREATE employee
exports.create = async (req, res) => {
  try {
    const { name, email, position, department, salary } = req.body;
    if (!name || !email)
      return res.status(400).json({ message: 'Name and Email are required' });

    const employee = new Employee({ name, email, position, department, salary });
    await employee.save();

    await updateExcelDataset(); // Update Excel after creation

    res.status(201).json(employee);
  } catch (err) {
    console.error(err);
    res.status(400).json({ message: err.message });
  }
};

// UPDATE employee
exports.update = async (req, res) => {
  try {
    const emp = await Employee.findByIdAndUpdate(req.params.id, req.body, { new: true });
    if (!emp) return res.status(404).json({ message: 'Employee not found' });

    await updateExcelDataset(); // Update Excel after update

    res.json(emp);
  } catch (err) {
    console.error(err);
    res.status(400).json({ message: err.message });
  }
};

// DELETE employee
exports.delete = async (req, res) => {
  try {
    const emp = await Employee.findByIdAndDelete(req.params.id);
    if (!emp) return res.status(404).json({ message: 'Employee not found' });

    await updateExcelDataset(); // Update Excel after delete

    res.json({ message: 'Employee deleted' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: err.message });
  }
};

// ------------------- EXCEL EXPORT -------------------

// Generate / Update Excel dataset
async function updateExcelDataset() {
  const employees = await Employee.find().lean();

  // Clean data for Excel (remove _id, __v)
  const excelData = employees.map(({ _id, __v, ...rest }) => rest);

  const worksheet = XLSX.utils.json_to_sheet(excelData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Employees');

  XLSX.writeFile(workbook, path.join(__dirname, '..', 'employee_dataset.xlsx'));
}

// Route: Download Excel file
exports.downloadExcel = async (req, res) => {
  try {
    const filePath = path.join(__dirname, '..', 'employee_dataset.xlsx');

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ message: 'Excel file not found' });
    }

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="employee_dataset.xlsx"'
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    fs.createReadStream(filePath).pipe(res);
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Failed to download Excel' });
  }
};