// ================= IMPORTS =================
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");
const nodemailer = require("nodemailer");
const crypto = require("crypto");
const dns = require("dns").promises;
const Member = require("./models/Member");
const cron = require("node-cron");
const Razorpay = require("razorpay");
const twilio = require("twilio");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const Complaint = require("./models/Complaint");

require("dotenv").config();


const client = twilio(
  process.env.TWILIO_SID,
  process.env.TWILIO_AUTH
);

const razorpay = new Razorpay({
  key_id: process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET
});

const app = express();
app.use(express.json());
app.use(cors());

const SECRET = process.env.JWT_SECRET;

// 🔥 AUTO REMINDER EVERY DAY AT 10 AM
cron.schedule("0 10 * * *", async () => {
  console.log("⏰ Running auto reminder...");

  const members = await Member.find();

  for (let m of members) {

    let profileDue = 0;

    (m.payments || []).forEach(p => {
      if (p.status !== "paid") {
        profileDue += p.amount;
      }
    });

    if (
  profileDue > 0 &&
  m.phone &&
  (!m.lastReminder || new Date(m.lastReminder).toDateString() !== today)
) {
  await sendWhatsApp(
    formatPhone(m.phone),
    `❗ Reminder

Hello ${m.name},

You have pending maintenance of ₹${profileDue}.

Please pay as soon as possible 🙏`
  );

  m.lastReminder = new Date();   // ✅ SAVE TIME
  await m.save();

  console.log(`📲 Sent to ${m.name}`);
}
  }

  console.log("✅ Auto reminder finished");
});

// ================= DB =================
mongoose.connect(process.env.MONGO_URI)
.then(()=>console.log("✅ DB Connected"))
.catch(err=>console.log(err));


// ================= AUTH =================
function auth(req,res,next){
  const token=req.headers.authorization;
  if(!token) return res.status(401).send("No token");

  try{
    req.user=jwt.verify(token,SECRET);
    next();
  }catch{
    res.status(401).send("Invalid token");
  }
}

function adminOnly(req,res,next){

  if(req.user.role !== "admin"){
    return res.status(403).send("Admin only");
  }

  next();
}


// ================= HELPER =================
function getType(p){
  if(p.status==="paid") return "paid";
  const diff = (new Date() - new Date(p.createdAt)) / (1000*60*60*24);
  return diff <= 5 ? "current" : "due";
}

function monthToNumber(month){
  return new Date(Date.parse(month +" 1, 2020")).getMonth() + 1;
}

function moneyToPaise(value){
  const amount = Number(value);
  if(!Number.isFinite(amount)) return 0;
  return Math.round(amount * 100);
}

function paiseToMoney(paise){
  return Math.round(Number(paise || 0)) / 100;
}

function getRazorpayMode(){
  const keyId = process.env.RAZORPAY_KEY_ID || "";

  if(keyId.startsWith("rzp_live_")) return "live";
  if(keyId.startsWith("rzp_test_")) return "test";

  return keyId ? "unknown" : "missing";
}

function normalizeEmail(email){
  return String(email || "").trim().toLowerCase();
}

function maskEmail(email){
  const cleanEmail = normalizeEmail(email);
  const [name, domain] = cleanEmail.split("@");

  if(!name || !domain) return "";

  return `${name.slice(0,2)}***@${domain}`;
}

function escapeRegex(value){
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function hashToken(token){
  return crypto.createHash("sha256").update(token).digest("hex");
}

async function createMailTransporter(){
  const [smtpHost] = await dns.resolve4("smtp.gmail.com");

  return nodemailer.createTransport({
    host: smtpHost || "smtp.gmail.com",
    port: 587,
    secure: false,
    requireTLS: true,
    connectionTimeout: 10000,
    greetingTimeout: 10000,
    socketTimeout: 10000,
    tls: {
      servername: "smtp.gmail.com"
    },
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
}

function shouldUseMailRelay(){
  return Boolean(process.env.RENDER || process.env.MAIL_RELAY_URL);
}

function getMailRelayUrl(){
  return process.env.MAIL_RELAY_URL || "https://newtownsociety-frontend.vercel.app/api/send-email";
}

function getMailRelaySecret(){
  return process.env.JWT_SECRET;
}

async function relayMail(mailOptions){
  const relaySecret = getMailRelaySecret();

  if(!relaySecret){
    throw new Error("Mail relay secret is not configured");
  }

  const response = await fetch(getMailRelayUrl(), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-mail-relay-secret": relaySecret
    },
    body: JSON.stringify({
      to: mailOptions.to,
      subject: mailOptions.subject,
      text: mailOptions.text
    })
  });

  if(!response.ok){
    const message = await response.text();
    throw new Error(message || "Mail relay failed");
  }
}

async function verifyMailRelay(){
  const relaySecret = getMailRelaySecret();

  if(!relaySecret){
    throw new Error("Mail relay secret is not configured");
  }

  const response = await fetch(getMailRelayUrl(), {
    headers: {
      "x-mail-relay-secret": relaySecret
    }
  });

  if(!response.ok){
    const message = await response.text();
    throw new Error(message || "Mail relay health check failed");
  }
}

function sendMailInBackground(mailOptions, context){
  (shouldUseMailRelay()
    ? relayMail(mailOptions)
    : createMailTransporter().then(transporter => transporter.sendMail(mailOptions)))
    .then(() => console.log(`${context} email sent`))
    .catch(error => console.log(`${context} email failed:`, error.message));
}

async function sendMailNow(mailOptions, context){
  try {
    if(shouldUseMailRelay()){
      await relayMail(mailOptions);
    }else{
      const transporter = await createMailTransporter();
      await transporter.sendMail(mailOptions);
    }

    console.log(`${context} email sent`);
  } catch (error) {
    console.log(`${context} email failed:`, error.message);
    throw new Error("Email could not be sent. Check EMAIL_USER and EMAIL_PASS in Render.");
  }
}

async function sendEmailVerificationOtp(user, options = {}){
  if(!user.email){
    throw new Error("No email registered");
  }

  const otp = Math.floor(100000 + Math.random() * 900000);

  user.emailVerificationOtp = otp;
  user.emailVerificationExpiry = Date.now() + 10 * 60 * 1000;

  await user.save();

  const mailOptions = {
    to: user.email,
    subject: "Verify your New Town Society account",
    text: `Hello ${user.name || "Resident"},

Your New Town Society email verification OTP is ${otp}.

This OTP will expire in 10 minutes.

If you did not try to login, please contact the society office.`
  };

  if(options.waitForDelivery){
    await sendMailNow(mailOptions, `Verification OTP for ${user.email}`);
    return;
  }

  sendMailInBackground(mailOptions, `Verification OTP for ${user.email}`);
}



// ================= LOGIN =================
app.post("/login", async(req,res)=>{
  
  const user=await Member.findOne({flatNumber:req.body.flatNumber});
  if(!user) return res.json({success:false});

  const match=await bcrypt.compare(req.body.password,user.password);
  if(!match) return res.json({success:false});

  if(user.role === "owner" && user.emailVerified !== true){
    if(!user.email){
      return res.json({
        success:false,
        verificationRequired:true,
        message:"Email verification required. No email is registered for this flat. Contact admin."
      });
    }

    try{
      await sendEmailVerificationOtp(user, { waitForDelivery: true });
    }catch(err){
      console.log("Email verification OTP error:", err.message);
      return res.status(500).json({
        success:false,
        verificationRequired:true,
        message:err.message || "Could not send verification OTP. Try again later."
      });
    }

    return res.json({
      success:false,
      verificationRequired:true,
      email: maskEmail(user.email),
      message:"Email verification required. OTP sent to registered email."
    });
  }

  const token=jwt.sign({flat:user.flatNumber,role:user.role},SECRET);
  res.json({success:true,token,role:user.role});
});

app.post("/verify-email", async(req,res)=>{
  const { flatNumber, otp } = req.body;

  const user = await Member.findOne({ flatNumber });

  if(!user) return res.status(404).json({ success:false, message:"User not found" });
  if(user.role !== "owner") return res.status(400).json({ success:false, message:"Only owners need email verification" });
  if(user.emailVerified === true) return res.json({ success:true, message:"Email already verified" });

  if(!user.emailVerificationOtp || !user.emailVerificationExpiry){
    return res.status(400).json({ success:false, message:"Verification OTP not found. Please login again to resend OTP." });
  }

  if(String(user.emailVerificationOtp) !== String(otp)){
    return res.status(400).json({ success:false, message:"Wrong OTP" });
  }

  if(Date.now() > user.emailVerificationExpiry){
    return res.status(400).json({ success:false, message:"OTP expired. Please login again to resend OTP." });
  }

  user.emailVerified = true;
  user.emailVerificationOtp = null;
  user.emailVerificationExpiry = null;

  await user.save();

  res.json({ success:true, message:"Email verified successfully" });
});

app.get("/fix-owner-password", auth, adminOnly, async(req,res)=>{

  const owner = await Member.findOne({ role: "owner" });

  if(!owner){
    return res.send("Owner not found");
  }

  owner.password = await bcrypt.hash("123456",10);

  await owner.save();

  res.send("Owner password fixed");
});

app.post("/send-reminder", auth, adminOnly, async (req, res) => {
  const { memberId } = req.body;

  const member = await Member.findById(memberId);

  if (!member) return res.send("Member not found");

  // calculate pending
  let profileDue = 0;

  (member.payments || []).forEach(p => {
    if (p.status !== "paid") {
      profileDue += p.amount;
    }
  });

  if (profileDue <= 0) {
    return res.send("No pending amount");
  }

  await sendWhatsApp(
    formatPhone(member.phone),
    `❗ Reminder

Hello ${member.name},

You have pending maintenance of ₹${profileDue}.

Please pay as soon as possible 🙏`
  );

  res.send("Reminder sent");
});

// ================= MEMBERS =================
app.get("/members", auth, adminOnly, async (req, res) => {
  const members = await Member.find().select(
    "-password -otp -otpExpiry -emailVerificationOtp -emailVerificationExpiry -passwordResetTokenHash -passwordResetExpiry"
  );

  const updatedMembers = members.map(m => {
    let profileDue = 0;
    let totalPaid = 0;

    (m.payments || []).forEach(p => {
      if (p.status === "pending") {
        profileDue += p.amount;
      } else if (p.status === "paid") {
        totalPaid += p.amount;
      }
    });

    return {
      ...m._doc,
      pendingAmount: profileDue
    };
  });

  res.json(updatedMembers);
});

app.put("/member/:id", auth, adminOnly, async (req, res) => {
  const existingMember = await Member.findById(req.params.id);
  if(!existingMember) return res.status(404).send("Member not found");

  const updateData = {
  name: req.body.name,
  phone: req.body.phone,
  email: normalizeEmail(req.body.email),
  flatNumber: req.body.flatNumber   // ✅ ADD THIS LINE
};

  if(normalizeEmail(existingMember.email) !== normalizeEmail(req.body.email)){
    updateData.emailVerified = false;
    updateData.emailVerificationOtp = null;
    updateData.emailVerificationExpiry = null;
  }

  // 🔥 ONLY FOR MEMBERS (who have area)
  if (req.body.area && !isNaN(req.body.area)) {
    updateData.area = Number(req.body.area);
    updateData.monthlyMaintenance = Number(req.body.area) * 1.5;
  }

  await Member.findByIdAndUpdate(req.params.id, updateData);

  res.send("Updated");
});

app.delete("/member/:id", auth, adminOnly, async(req,res)=>{
  await Member.findByIdAndDelete(req.params.id);
  res.send("Deleted");
});

app.put("/update-profile", auth, async (req, res) => {
  const { userId, phone, email } = req.body;

  await Member.findByIdAndUpdate(userId, {
    phone,
    email
  });

  res.send("Profile updated");
});

async function sendPasswordResetLink(req, res){
  try {
    const { email } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail) return res.status(400).send("Enter email");

    const user = await Member.findOne({
      email: cleanEmail
    });

    if (!user) return res.status(404).send("No account found with this email");

    const resetToken = crypto.randomBytes(32).toString("hex");
    const appOrigin =
      req.get("origin") ||
      process.env.FRONTEND_URL ||
      "https://newtownsociety-frontend.vercel.app";
    const resetLink = `${appOrigin}/?resetToken=${encodeURIComponent(resetToken)}&email=${encodeURIComponent(cleanEmail)}`;

    user.passwordResetTokenHash = hashToken(resetToken);
    user.passwordResetExpiry = Date.now() + 15 * 60 * 1000;

    await user.save();

    console.log("Password reset link saved for:", cleanEmail);

    await sendMailNow({
      to: cleanEmail,
      subject: "Reset your New Town Society password",
      text: `Hello ${user.name || "Resident"},

Open this secure link to reset your New Town Society password:
${resetLink}

This link expires in 15 minutes. If you did not request this, please ignore this email.`
    }, `Password reset link for ${cleanEmail}`);

    res.send("Password reset link sent");
  } catch (error) {
    console.log("Send reset link error:", error);
    res.status(500).send(error.message || "Could not send reset link");
  }
}

app.post("/send-reset-link", sendPasswordResetLink);

app.post("/send-otp", sendPasswordResetLink);

app.post("/reset-password", async (req, res) => {
  try {
    const { email, token, newPassword } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail || !token || !newPassword) {
      return res.status(400).send("Open the reset link and enter a new password");
    }

    if (String(newPassword).length < 4) {
      return res.status(400).send("Password must be at least 4 characters");
    }

    const user = await Member.findOne({
      email: cleanEmail,
      passwordResetTokenHash: hashToken(token)
    });

    if (!user) return res.status(400).send("Reset link is invalid");

    if (!user.passwordResetExpiry || Date.now() > user.passwordResetExpiry) {
      return res.status(400).send("Reset link expired. Please request a new link");
    }

    user.password = await bcrypt.hash(newPassword, 10);
    user.passwordResetTokenHash = null;
    user.passwordResetExpiry = null;
    user.otp = null;
    user.otpExpiry = null;

    await user.save();

    res.send("Password updated");
  } catch (error) {
    console.log("Reset password error:", error);
    res.status(500).send("Could not reset password");
  }
});

app.post("/verify-otp", async (req, res) => {
  try {
    const { email, otp, newPassword } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail || !otp || !newPassword) {
      return res.status(400).send("Enter email, OTP, and new password");
    }

    const user = await Member.findOne({
      email: cleanEmail
    });

    if (!user) return res.status(404).send("No account found with this email");

    if (!user.otp || !user.otpExpiry) {
      return res.status(400).send("Please request a new OTP");
    }

    if (String(user.otp) !== String(otp)) {
      return res.status(400).send("Wrong OTP");
    }

    if (Date.now() > user.otpExpiry) {
      return res.status(400).send("OTP expired");
    }

    user.password = await bcrypt.hash(newPassword, 10);
    user.otp = null;
    user.otpExpiry = null;

    await user.save();

    res.send("Password updated");
  } catch (error) {
    console.log("Verify OTP error:", error);
    res.status(500).send("Could not reset password");
  }
});

// ================= ADD DUE =================
app.post("/add-due", auth, adminOnly, async(req,res)=>{
  const {memberId,month,year,amount}=req.body;

  const m = await Member.findById(memberId);

  m.payments.push({
  month,
  year,
  amount,
  status: "pending",   // 🔥 ADD THIS
  createdAt: new Date(2020,0,1)
});

  await m.save();

  await sendWhatsApp(
  formatPhone(m.phone),
  `❗ New Due Added

📅 ${month} ${year}
💰 Amount: ₹${amount}

Please pay on time 🙏`
);

  res.send("Due added");
});

app.post("/add-member", auth,adminOnly, async (req, res) => {
  try {
    const { name, flatNumber, phone, email, area } = req.body;

    const cleanName = String(name || "").trim();
    const cleanFlatNumber = String(flatNumber || "").trim();
    const cleanPhone = String(phone || "").trim();
    const cleanEmail = normalizeEmail(email);
    const cleanArea = Number(area);

    if(!cleanName || !cleanFlatNumber || !cleanPhone || !cleanEmail || !cleanArea){
      return res.status(400).send("Please fill all member details");
    }

    if(!/^\S+@\S+\.\S+$/.test(cleanEmail)){
      return res.status(400).send("Please enter a valid email address");
    }

    const existingMember = await Member.findOne({
      $or: [
        { flatNumber: { $regex: `^${escapeRegex(cleanFlatNumber)}$`, $options: "i" } },
        { email: cleanEmail }
      ]
    });

    if(existingMember){
      if(existingMember.flatNumber?.toLowerCase() === cleanFlatNumber.toLowerCase()){
        return res.status(409).send("This flat number already exists");
      }

      return res.status(409).send("This email is already used by another member");
    }

const hashedPassword = await bcrypt.hash(cleanFlatNumber, 10);

const newMember = new Member({
  name: cleanName,
  flatNumber: cleanFlatNumber,
  area: cleanArea,
  monthlyMaintenance: cleanArea * 1.5,
  phone: cleanPhone,
  email: cleanEmail,
  emailVerified: false,
  password: hashedPassword,
  role: "owner",
  payments: []
});

    await newMember.save();

    sendEmailVerificationOtp(newMember)
      .then(() => console.log(`Verification OTP sent to ${newMember.email}`))
      .catch(mailErr => console.log("Verification email failed:", mailErr.message));

    res.status(201).send("Member added successfully. Verification OTP is being sent to owner email.");
  } catch (err) {
    console.log("Add member error:", err);
    res.status(500).send(err.message || "Error adding member");
  }
});

app.get("/all-payments", auth, adminOnly, async (req, res) => {
  const members = await Member.find();

  let list = [];

  members.forEach(m => {
    (m.payments || []).forEach(p => {
      list.push({
        name: m.name,
        flat: m.flatNumber,
        month: p.month,
        year: p.year,
        amount: p.amount,
        status: p.status || "pending"
      });
    });
  });

  res.json(list);
});

app.post("/export-history", auth, adminOnly, async (req, res) => {

  try {
    const data = req.body;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Payments");

    ws.columns = [
      { header: "Name", key: "name", width: 20 },
      { header: "Flat", key: "flat", width: 10 },
      { header: "Month", key: "month", width: 15 },
      { header: "Year", key: "year", width: 10 },
      { header: "Amount", key: "amount", width: 15 },
      { header: "Status", key: "status", width: 15 }
    ];

    data.forEach(p => {
      ws.addRow({
        name: p.name,
        flat: p.flat,
        month: p.month,
        year: p.year,
        amount: p.amount,
        status: p.status
      });
    });

    // 🔥 VERY IMPORTANT HEADERS
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=history.xlsx"
    );

    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.log("❌ Excel error:", err);
    res.status(500).send("Excel failed");
  }
});

// ================= OWNER DASH =================
app.get("/my-dashboard", auth, async(req,res)=>{
  const user = await Member.findOne({flatNumber:req.user.flat});

  const now = new Date();
  const month = now.toLocaleString("default",{month:"long"});
  const year = now.getFullYear();

  // 🔥 CHECK IF CURRENT MONTH BILL EXISTS
  let exists = user.payments.find(p=>p.month===month && p.year===year);

  if(!exists){
    user.payments.push({
      month,
      year,
      amount: user.monthlyMaintenance
    });

    await user.save();
  }

  let current=[],due=[],paid=[];

  user.payments.forEach(p=>{
    if(p.status==="paid"){
      paid.push(p);
    }else{
      const diff = (new Date() - new Date(p.createdAt))/(1000*60*60*24);

      if(diff <= 5 && p.month===month && p.year===year){
        current.push(p);
      }else{
        due.push(p);
      }
    }
  });

  res.json({
  _id: user._id,
  name: user.name,
  flatNumber: user.flatNumber,
  phone: user.phone,
  email: user.email,
  current,
  due,
  paid
});
});


// ================= PAY (FIXED WITH ID) =================
app.post("/pay-now", auth, adminOnly, async(req,res)=>{
  const {paymentId} = req.body;

  const user = await Member.findOne({flatNumber:req.user.flat});
  const payment = user.payments.id(paymentId);

  payment.status="paid";
  payment.paidDate=new Date();

  await user.save();

  res.send("Paid");
});


// ================= PENDING =================
app.get("/pending", auth, adminOnly, async(req,res)=>{
  const members=await Member.find();

  let list=[];

  members.forEach(m=>{
    m.payments.forEach(p=>{
      if(p.status==="pending"){
        list.push({
          name:m.name,
          flat:m.flatNumber,
          amount:p.amount,
          month:p.month
        });
      }
    });
  });

  res.json(list);
});


// ================= EXCEL =================
app.get("/export", auth, adminOnly, async (req, res) => {
  const { from, to, flat } = req.query;

  console.log("FLAT RECEIVED:", flat);

  // ✅ DATE FIX
  const start = from ? new Date(from) : new Date("2000-01");
  const end = to ? new Date(to) : new Date();

  let members;

  // ✅ FLAT FILTER
  if (flat && flat.trim() !== "") {
    members = await Member.find({
  flatNumber: { $regex: flat.trim(), $options: "i" },
  role: "owner"   // 👈 FILTER ONLY OWNERS
});
  } else {
    members = await Member.find();
  }

  console.log("MEMBERS COUNT:", members.length);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Report");

  ws.columns = [
    { header: "Name", key: "name" },
    { header: "Flat", key: "flat" },
    { header: "Month", key: "month" },
    { header: "Year", key: "year" },
    { header: "Amount", key: "amount" },
    { header: "Status", key: "status" }
  ];

  members.forEach(m => {
    m.payments.forEach(p => {

      const pDate = new Date(
        p.year,
        monthToNumber(p.month) - 1,
        1
      );

      if (pDate >= start && pDate <= end) {
        ws.addRow({
          name: m.name,
          flat: m.flatNumber,
          month: p.month,
          year: p.year,
          amount: p.amount,
          status: p.status
        });
      }

    });
  });

  res.setHeader(
    "Content-Disposition",
    "attachment; filename=report.xlsx"
  );

  await wb.xlsx.write(res);
  res.end();
});

// 🔥 AUTO MONTHLY BILL (PRO VERSION)
cron.schedule("0 0 1 * *", async () => {
  try {
    console.log("📅 Running monthly maintenance...");

    const now = new Date();
    const currentMonth = now.toLocaleString("default", { month: "long" });
    const currentYear = now.getFullYear();

    const members = await Member.find();

    for (let m of members) {

      // ❗ Prevent duplicate entry
      const alreadyExists = m.payments.find(p =>
        p.month === currentMonth && p.year === currentYear
      );

      if (!alreadyExists) {
        m.payments.push({
          month: currentMonth,
          year: currentYear,
          amount: m.monthlyMaintenance,
          status: "pending",
          createdAt: new Date()   // 🔥 added
        });

        await m.save();

        await sendWhatsApp(
  formatPhone(m.phone),
  `📅 Monthly Maintenance Added

💰 ₹${m.monthlyMaintenance}

Please pay within 5 days 🙏`
);

        console.log(`✅ Bill added for ${m.name}`);
      } else {
        console.log(`⚠️ Already exists for ${m.name}`);
      }
    }

    console.log("🎉 Monthly billing completed");

  } catch (err) {
    console.log("❌ Cron Error:", err);
  }
});

app.post("/create-order", auth, async (req, res) => {
  const { amount, paymentId } = req.body;
  const amountPaise = moneyToPaise(amount);

  if (amountPaise <= 0) {
    return res.status(400).json({ error: "Enter a valid amount" });
  }

  if (!paymentId) {
    return res.status(400).json({ error: "Payment is required" });
  }

  const member = await Member.findOne({
    flatNumber: req.user.flat,
    "payments._id": paymentId
  });

  if (!member) {
    return res.status(404).json({ error: "Payment not found" });
  }

  const payment = member.payments.find(p => p._id.toString() === paymentId);

  if (!payment || payment.status === "paid") {
    return res.status(400).json({ error: "This due is already paid" });
  }

  const pendingTotalPaise = member.payments
    .filter(p => p.status !== "paid")
    .reduce((sum, p) => sum + moneyToPaise(p.amount), 0);

  if (amountPaise > pendingTotalPaise) {
    return res.status(400).json({ error: "Amount cannot exceed pending due" });
  }

  const options = {
    amount: amountPaise,
    currency: "INR",
    receipt: "receipt_" + Date.now()
  };

  try {
    const order = await razorpay.orders.create(options);
    res.json(order);
  } catch (err) {
    console.log(err);

    // 🔥 FIXED LINE
    res.status(500).json({ error: "Order creation failed" });
  }
});

app.get("/payment-config", auth, (req, res) => {
  const keyId = process.env.RAZORPAY_KEY_ID || "";

  res.json({
    razorpayKeyId: keyId,
    mode: getRazorpayMode()
  });
});

app.post("/verify-payment", async (req, res) => {
  try {
    console.log("🔥 BODY RECEIVED:", req.body);

    const {
      paymentId,
      razorpay_payment_id,
      razorpay_order_id,
      razorpay_signature,
      paidAmount
    } = req.body;

    // ❌ check missing fields
    if (!paymentId || !razorpay_payment_id || !razorpay_order_id || !razorpay_signature) {
      console.log("❌ Missing payment data");
      return res.status(400).json({ error: "Invalid payment data" });
    }

    // ✅ VERIFY SIGNATURE
    if (!process.env.RAZORPAY_KEY_SECRET) {
      console.log("❌ Missing RAZORPAY_KEY_SECRET");
      return res.status(500).json({ error: "Payment secret is not configured" });
    }

    const generated_signature = crypto
      .createHmac("sha256", process.env.RAZORPAY_KEY_SECRET)
      .update(razorpay_order_id + "|" + razorpay_payment_id)
      .digest("hex");

    console.log("Generated:", generated_signature);
    console.log("Received:", razorpay_signature);

    if (generated_signature !== razorpay_signature) {
      console.log("❌ Signature mismatch");
      return res.status(400).json({ error: "Payment verification failed" });
    }

    console.log("✅ Signature verified");

    // 🔍 find member
    const member = await Member.findOne({
      "payments._id": paymentId
    });

    if (!member) {
      console.log("❌ Member not found");
      return res.status(404).json({ error: "Member not found" });
    }

    // 🔍 find payment
    const payment = member.payments.find(
      p => p._id.toString() === paymentId
    );

    if (!payment) {
      console.log("❌ Payment not found");
      return res.status(404).json({ error: "Payment not found" });
    }

    if (payment.status === "paid") {
      console.log("❌ Payment already paid");
      return res.status(400).json({ error: "This due is already paid" });
    }

    const amountPaise = moneyToPaise(paidAmount ?? payment.amount);

    if (amountPaise <= 0) {
      return res.status(400).json({ error: "Enter a valid amount" });
    }

    const pendingPayments = member.payments.filter(
      p => p.status !== "paid" && moneyToPaise(p.amount) > 0
    );
    const pendingTotalPaise = pendingPayments.reduce(
      (sum, p) => sum + moneyToPaise(p.amount),
      0
    );

    if (amountPaise > pendingTotalPaise) {
      return res.status(400).json({ error: "Amount cannot exceed pending due" });
    }

    let order;
    try {
      order = await razorpay.orders.fetch(razorpay_order_id);
    } catch (error) {
      console.log("❌ Razorpay order fetch failed:", error.message);
      return res.status(400).json({ error: "Payment order could not be verified" });
    }

    if (Number(order.amount) !== amountPaise) {
      console.log("❌ Amount mismatch:", order.amount, amountPaise);
      return res.status(400).json({ error: "Payment amount mismatch" });
    }

    const selectedPayment = pendingPayments.find(
      p => p._id.toString() === paymentId
    );
    const orderedPendingPayments = [
      selectedPayment,
      ...pendingPayments.filter(p => p._id.toString() !== paymentId)
    ].filter(Boolean);

    let remainingPaise = amountPaise;
    const paidDate = new Date();
    const appliedPayments = [];

    for (const pendingPayment of orderedPendingPayments) {
      if (remainingPaise <= 0) break;

      const duePaise = moneyToPaise(pendingPayment.amount);
      if (duePaise <= 0) continue;

      const appliedPaise = Math.min(duePaise, remainingPaise);
      const appliedAmount = paiseToMoney(appliedPaise);

      if (appliedPaise >= duePaise) {
        pendingPayment.status = "paid";
        pendingPayment.paidDate = paidDate;
        appliedPayments.push({
          _id: pendingPayment._id,
          month: pendingPayment.month,
          year: pendingPayment.year,
          amount: appliedAmount
        });
      } else {
        pendingPayment.amount = paiseToMoney(duePaise - appliedPaise);
        member.payments.push({
          month: pendingPayment.month,
          year: pendingPayment.year,
          amount: appliedAmount,
          status: "paid",
          paidDate
        });
        appliedPayments.push(member.payments[member.payments.length - 1]);
      }

      remainingPaise -= appliedPaise;
    }

    if (remainingPaise > 0) {
      return res.status(400).json({ error: "Amount cannot exceed pending due" });
    }

    await member.save();

    console.log("✅ DB UPDATED");

    const paidTotal = paiseToMoney(amountPaise);
    const receiptPayment = appliedPayments.length === 1
      ? appliedPayments[0]
      : {
          _id: razorpay_payment_id,
          month: "Multiple dues",
          year: new Date().getFullYear(),
          amount: paidTotal
        };
    const remainingDue = member.payments.reduce(
      (sum, p) => p.status !== "paid" ? sum + Number(p.amount || 0) : sum,
      0
    );

    // 📄 generate bill
    const billFile = await generateBill(member, receiptPayment);

    // 📲 send whatsapp
    await sendWhatsApp(
      formatPhone(member.phone),
      `✅ Payment Received!

📅 Month: ${receiptPayment.month}
💰 Amount: ₹${paidTotal}
⏳ Remaining Due: ₹${remainingDue}

📄 Download Bill:
http://localhost:5000/bills/${billFile}

Thank you 🙏`
    );

    console.log("✅ WhatsApp sent");

    res.json({ success: true, paidAmount: paidTotal, remainingDue });

  } catch (err) {
    console.log("❌ ERROR:", err);
    res.status(500).json({ error: "Verification failed" });
  }
});

async function sendWhatsApp(to, message) {
  try {
    const msg = await client.messages.create({
      from: "whatsapp:+14155238886",
      to: "whatsapp:" + to,
      body: message
    });

    console.log("✅ Message sent:", msg.sid);
  } catch (err) {
    console.log("❌ Error:", err.message);
  }
}

app.get("/test-whatsapp", async (req, res) => {

  await sendWhatsApp(
    "+918670433655",  // 👈 PUT YOUR NUMBER
    "🔥 WhatsApp working from your app!"
  );

  res.send("Message sent");
});

function formatPhone(phone) {
  phone = phone.replace(/\s+/g, "");

  if (phone.startsWith("+")) return phone;

  if (phone.startsWith("0")) {
    phone = phone.substring(1);
  }

  return "+91" + phone;
}

app.use("/bills", express.static("bills"));

function generateBill(member, payment) {
  return new Promise((resolve, reject) => {

    const fileName = `bill_${payment._id}.pdf`;
    const filePath = `./bills/${fileName}`;

    const doc = new PDFDocument();
    const stream = fs.createWriteStream(filePath);

    doc.pipe(stream);

    doc.fontSize(20).text("Society Maintenance Bill", { align: "center" });
    doc.moveDown();

    doc.text(`Name: ${member.name}`);
    doc.text(`Flat: ${member.flatNumber}`);
    doc.text(`Phone: ${member.phone}`);

    doc.moveDown();

    doc.text(`Month: ${payment.month}`);
    doc.text(`Year: ${payment.year}`);
    doc.text(`Amount: ₹${payment.amount}`);
    doc.text(`Payment ID: ${payment._id}`);
    doc.text(`Date: ${new Date().toLocaleString()}`);

    doc.moveDown();
    doc.text("Thank you 🙏");

    doc.end();

    stream.on("finish", () => resolve(fileName));
    stream.on("error", reject);
  });
}

// ================= COMPLAINT =================

// ADD COMPLAINT
app.post("/add-complaint", auth, async (req,res)=>{

  try{

    const member = await Member.findOne({
      flatNumber: req.user.flat
    });

    if(!member){
      return res.send("Member not found");
    }

    await Complaint.create({
      name: member.name,
      flat: member.flatNumber,
      phone: member.phone,
      message: req.body.message
    });

    res.send("Complaint submitted ✅");

  }catch(err){
    console.log(err);
    res.status(500).send("Error");
  }
});

// GET ALL COMPLAINTS
app.get("/complaints", auth, adminOnly, async(req,res)=>{

  try{

    const complaints = await Complaint
      .find()
      .sort({ date: -1 });

    res.json(complaints);

  }catch(err){
    console.log(err);
    res.status(500).send([]);
  }
});

app.get("/health", (req, res) => {
  res.json({
    ok: true,
    build: "vercel-mail-relay-jwt-secret",
    emailUser: maskEmail(process.env.EMAIL_USER),
    emailPassConfigured: Boolean(process.env.EMAIL_PASS),
    mailRelay: shouldUseMailRelay() ? getMailRelayUrl() : null,
    razorpayConfigured: Boolean(process.env.RAZORPAY_KEY_ID && process.env.RAZORPAY_KEY_SECRET),
    razorpayMode: getRazorpayMode()
  });
});

app.get("/health-email", async (req, res) => {
  try {
    if(shouldUseMailRelay()){
      await verifyMailRelay();
    }else{
      const transporter = await createMailTransporter();
      await transporter.verify();
    }

    res.json({
      ok: true,
      emailUser: maskEmail(process.env.EMAIL_USER),
      via: shouldUseMailRelay() ? "relay" : "smtp"
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      emailUser: maskEmail(process.env.EMAIL_USER),
      via: shouldUseMailRelay() ? "relay" : "smtp",
      code: error.code,
      responseCode: error.responseCode,
      command: error.command,
      message: error.message
    });
  }
});

// ================= SERVER =================
const PORT = process.env.PORT || 5000;

app.listen(PORT, () =>
  console.log(`🚀 Server running on ${PORT}`)
);
