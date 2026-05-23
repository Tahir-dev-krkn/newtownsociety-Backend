// ================= IMPORTS =================
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");
const nodemailer = require("nodemailer");
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

const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 465,
  secure: true,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  }
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

function normalizeEmail(email){
  return String(email || "").trim().toLowerCase();
}

function maskEmail(email){
  const cleanEmail = normalizeEmail(email);
  const [name, domain] = cleanEmail.split("@");

  if(!name || !domain) return "";

  return `${name.slice(0,2)}***@${domain}`;
}

async function sendEmailVerificationOtp(user){
  if(!user.email){
    throw new Error("No email registered");
  }

  const otp = Math.floor(100000 + Math.random() * 900000);

  user.emailVerificationOtp = otp;
  user.emailVerificationExpiry = Date.now() + 10 * 60 * 1000;

  await user.save();

  await transporter.sendMail({
    to: user.email,
    subject: "Verify your New Town Society account",
    text: `Hello ${user.name || "Resident"},

Your New Town Society email verification OTP is ${otp}.

This OTP will expire in 10 minutes.

If you did not try to login, please contact the society office.`
  });
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
      await sendEmailVerificationOtp(user);
    }catch(err){
      console.log("Email verification OTP error:", err.message);
      return res.status(500).json({
        success:false,
        verificationRequired:true,
        message:"Could not send verification OTP. Try again later."
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
app.get("/members", auth, async (req, res) => {
  const members = await Member.find();

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

app.post("/send-otp", async (req, res) => {
  const { email } = req.body;

  const user = await Member.findOne({
    email: normalizeEmail(email)
  });

  if (!user) return res.send("User not found");

  const otp = Math.floor(100000 + Math.random() * 900000);

  user.otp = otp;
  user.otpExpiry = Date.now() + 5 * 60 * 1000;

  await user.save(); // 🔥 THIS IS THE REAL FIX

  console.log("✅ OTP SAVED:", user.otp);

  await transporter.sendMail({
    to: email,
    subject: "OTP Verification",
    text: `Your OTP is ${otp}`
  });

  res.send("OTP sent");
});

app.post("/verify-otp", async (req, res) => {
  const { email, otp, newPassword } = req.body;

  const user = await Member.findOne({
    email: normalizeEmail(email)
  });

  if (!user) return res.send("User not found");

  console.log("ENTERED OTP:", otp);
  console.log("DB OTP:", user.otp);

  // 🔥 FORCE STRING MATCH (MOST IMPORTANT)
  if (String(user.otp) !== String(otp)) {
    return res.send("Wrong OTP");
  }

  if (Date.now() > user.otpExpiry) {
    return res.send("OTP expired");
  }

  user.password = await bcrypt.hash(newPassword, 10);
  user.otp = null;
  user.otpExpiry = null;

  await user.save();

  res.send("Password updated");
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

const hashedPassword = await bcrypt.hash(flatNumber, 10);

const newMember = new Member({
  name,
  flatNumber,
  area: Number(area),
  monthlyMaintenance: Number(area) * 1.5,
  phone,
  email: normalizeEmail(email),
  emailVerified: false,
  password: hashedPassword,
  role: "owner",
  payments: []
});

    await newMember.save();

    try{
      await sendEmailVerificationOtp(newMember);
      res.send("Member added successfully. Verification OTP sent to owner email.");
    }catch(mailErr){
      console.log("Verification email failed:", mailErr.message);
      res.send("Member added successfully, but verification email could not be sent.");
    }
  } catch (err) {
    console.log(err);
    res.status(500).send("Error adding member");
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
app.get("/pending", auth, async(req,res)=>{
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

app.post("/create-order", async (req, res) => {
  const { amount } = req.body;

  const options = {
    amount: amount * 100,
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

const crypto = require("crypto");

app.post("/verify-payment", async (req, res) => {
  try {
    console.log("🔥 BODY RECEIVED:", req.body);

    const { 
      paymentId, 
      razorpay_payment_id, 
      razorpay_order_id, 
      razorpay_signature 
    } = req.body;

    // ❌ check missing fields
    if (!razorpay_payment_id || !razorpay_order_id || !razorpay_signature) {
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

    // ✅ UPDATE DB
    await Member.updateOne(
      { "payments._id": paymentId },
      {
        $set: {
          "payments.$.status": "paid",
          "payments.$.paidDate": new Date()
        }
      }
    );

    console.log("✅ DB UPDATED");

    // 🔄 GET UPDATED DATA
    const updatedMember = await Member.findOne({
      "payments._id": paymentId
    });

    const updatedPayment = updatedMember.payments.find(
      p => p._id.toString() === paymentId
    );

    // 📄 generate bill
    const billFile = await generateBill(updatedMember, updatedPayment);

    // 📲 send whatsapp
    await sendWhatsApp(
      formatPhone(updatedMember.phone),
      `✅ Payment Received!

📅 Month: ${updatedPayment.month}
💰 Amount: ₹${updatedPayment.amount}

📄 Download Bill:
http://localhost:5000/bills/${billFile}

Thank you 🙏`
    );

    console.log("✅ WhatsApp sent");

    res.json({ success: true });

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

// ================= SERVER =================
const PORT = process.env.PORT || 5000;

app.listen(PORT, () =>
  console.log(`🚀 Server running on ${PORT}`)
);
