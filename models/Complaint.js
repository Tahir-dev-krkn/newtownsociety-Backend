const mongoose = require("mongoose");

const ComplaintSchema = new mongoose.Schema({
  name: String,
  flat: String,
  phone: String,
  message: String,
  date: {
    type: Date,
    default: Date.now
  }
});

module.exports = mongoose.model(
  "Complaint",
  ComplaintSchema
);