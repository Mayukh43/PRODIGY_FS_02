const express = require('express');
const {
  getAll,
  getOne,           // added
  create,
  update,
  delete: remove,
  downloadExcel
} = require('../controllers/employeeController');
const verifyToken = require('../middleware/authMiddleware');

const router = express.Router();

// CRUD routes
router.get('/', verifyToken, getAll);
router.get('/:id', verifyToken, getOne); // fetch single employee
router.post('/', verifyToken, create);
router.put('/:id', verifyToken, update);
router.delete('/:id', verifyToken, remove);

// Excel download route
router.get('/download-excel', verifyToken, downloadExcel);

module.exports = router;