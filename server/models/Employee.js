const mongoose = require('mongoose');

const employeeSchema = new mongoose.Schema({
  name: { 
    type: String, 
    required: true 
  },
  email: {            // added email
    type: String,
    required: true,
    unique: true
  },
  position: String,
  department: String,
  salary: Number,
  hireDate: {
    type: Date,
    default: Date.now,
  }
}, { timestamps: true }); // adds createdAt and updatedAt

module.exports = mongoose.model('Employee', employeeSchema);
