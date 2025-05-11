const express = require('express');
const Task = require('../models/Task');
const auth = require('../middleware/auth');
const excelJS = require('exceljs');
const multer = require('multer');
const Papa = require('papaparse');

const router = express.Router();
const upload = multer();

// Create a task
router.post('/', auth, async (req, res) => {
    try {
        const task = new Task({ ...req.body, user: req.user.id });
        await task.save();
        res.status(201).json(task);
    } catch (err) {
        res.status(400).json(err);
    }
});

// Get all tasks for logged-in user
router.get('/', auth, async (req, res) => {
    try {
        const tasks = await Task.find({ user: req.user.id });
        res.json(tasks);
    } catch (err) {
        res.status(500).json(err);
    }
});

// Update a task
router.put('/:id', auth, async (req, res) => {
    try {
        const task = await Task.findOneAndUpdate(
            { _id: req.params.id, user: req.user.id },
            req.body,
            { new: true }
        );
        res.json(task);
    } catch (err) {
        res.status(500).json(err);
    }
});

// Delete a task
router.delete('/:id', auth, async (req, res) => {
    try {
        await Task.findOneAndDelete({ _id: req.params.id, user: req.user.id });
        res.status(204).send();
    } catch (err) {
        res.status(500).json(err);
    }
});

// Export tasks to Excel
router.get('/export', auth, async (req, res) => {
    try {
        const tasks = await Task.find({ user: req.user.id });

        const workbook = new excelJS.Workbook();
        const worksheet = workbook.addWorksheet('Tasks');

        worksheet.columns = [
            { header: 'Title', key: 'title', width: 30 },
            { header: 'Description', key: 'description', width: 30 },
            { header: 'Effort (Days)', key: 'effort', width: 15 },
            { header: 'Due Date', key: 'dueDate', width: 20 },
        ];

        tasks.forEach(task => {
            worksheet.addRow({
                title: task.title,
                description: task.description,
                effort: task.effort,
                dueDate: task.dueDate,
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=tasks.xlsx');

        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        res.status(500).json(err);
    }
});

// Upload bulk tasks (CSV)
router.post('/upload', auth, upload.single('file'), async (req, res) => {
    try {
        const parsed = Papa.parse(req.file.buffer.toString(), { header: true });
        const tasks = parsed.data.map(t => ({
            ...t,
            user: req.user.id,
            effort: Number(t.effort),
            dueDate: new Date(t.dueDate),
        }));

        await Task.insertMany(tasks);
        res.status(201).send('Tasks uploaded');
    } catch (err) {
        res.status(500).json({ msg: 'Failed to upload tasks', error: err });
    }
});

module.exports = router;
