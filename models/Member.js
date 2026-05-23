const mongoose = require("mongoose");

const MemberSchema = new mongoose.Schema({
  name: String,
  flatNumber: String,
  area: Number,
  phone: String,
  email: String,
  emailVerified: {
    type: Boolean,
    default: false
  },
  emailVerificationOtp: Number,
  emailVerificationExpiry: Number,
  password: String,
  role: String,
  monthlyMaintenance: Number,
  lastReminder: Date,

  // 🔥 OTP FIELDS
  otp: Number,
  otpExpiry: Number,
  passwordResetTokenHash: String,
  passwordResetExpiry: Number,

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

module.exports = mongoose.model("Member", MemberSchema);
