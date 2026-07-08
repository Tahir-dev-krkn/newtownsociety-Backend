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
      description: String,
      chargeType: {
        type: String,
        default: "maintenance"
      },
      paymentMethod: String,
      razorpayOrderId: String,
      razorpayPaymentId: String,
      status: {
        type: String,
        default: "pending"
      },
      createdAt: {
        type: Date,
        default: Date.now
      },
      paidDate: Date
    }
  ]
});

// 🔎 Indexes for fast login / lookup (avoids full-collection scans)
MemberSchema.index({ flatNumber: 1 });
MemberSchema.index({ email: 1 });
MemberSchema.index({ "payments._id": 1 });

module.exports = mongoose.model("Member", MemberSchema);
