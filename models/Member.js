const mongoose = require("mongoose");

const MemberSchema = new mongoose.Schema({
  name: String,
  flatNumber: String,
  area: Number,
  phone: String,
  email: String,
  password: String,
  role: String,
  monthlyMaintenance: Number,
  lastReminder: Date,

  // 🔥 OTP FIELDS
  otp: Number,
  otpExpiry: Number,

  // 🔥 PAYMENTS (INSIDE SCHEMA)
  payments: [
    {
      month: String,
      year: Number,
      amount: Number,
      status: {
        type: String,
        default: "pending"
      },
      paidDate: Date
    }
  ]
});

complaints: [
  {
    message: String,
    date: {
      type: Date,
      default: Date.now
    }
  }
],



module.exports = mongoose.model("Member", MemberSchema);