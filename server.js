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
const rateLimit = require("express-rate-limit");

require("dotenv").config();


const TWILIO_ACCOUNT_SID = firstEnv("TWILIO_SID", "TWILIO_ACCOUNT_SID");
const TWILIO_AUTH_TOKEN = firstEnv("TWILIO_AUTH", "TWILIO_AUTH_TOKEN");
const client = TWILIO_ACCOUNT_SID && TWILIO_AUTH_TOKEN
  ? twilio(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
  : null;

const TWILIO_SANDBOX_FROM = "whatsapp:+14155238886";
const WHATSAPP_FROM = normalizeWhatsAppAddress(
  process.env.TWILIO_WHATSAPP_FROM,
  TWILIO_SANDBOX_FROM
);
const TWILIO_MESSAGING_SERVICE_SID = firstEnv(
  "TWILIO_MESSAGING_SERVICE_SID",
  "TWILIO_MESSAGE_SERVICE_SID",
  "TWILIO_MESSAGING_SID"
);
const BACKEND_PUBLIC_URL = String(
  process.env.BACKEND_PUBLIC_URL ||
  process.env.PUBLIC_BACKEND_URL ||
  process.env.RENDER_EXTERNAL_URL ||
  ""
).replace(/\/+$/, "");
const APP_PUBLIC_URL = String(
  process.env.APP_PUBLIC_URL ||
  process.env.FRONTEND_PUBLIC_URL ||
  process.env.NEXT_PUBLIC_APP_URL ||
  "https://newtwnsociety.com"
).replace(/\/+$/, "");
const TWILIO_TEMPLATES = {
  reminder: firstEnv(
    "TWILIO_TEMPLATE_REMINDER_SID",
    "TWILIO_TEMPLATE_BALANCE_DUE_SID",
    "TWILIO_TEMPLATE_BALANCE_UPDATE_SID",
    "TWILIO_TEMPLATE_MAINTENANCE_REMINDER_SID"
  ),
  dueAdded: firstEnv(
    "TWILIO_TEMPLATE_DUE_ADDED_SID",
    "TWILIO_TEMPLATE_CHARGE_ADDED_SID",
    "TWILIO_TEMPLATE_DUE_UPDATE_SID",
    "TWILIO_TEMPLATE_CHARGE_NOTICE_SID"
  ),
  monthlyDue: firstEnv(
    "TWILIO_TEMPLATE_MONTHLY_DUE_SID",
    "TWILIO_TEMPLATE_MONTHLY_BILL_SID",
    "TWILIO_TEMPLATE_MONTHLY_UPDATE_SID"
  ),
  paymentReceived: firstEnv(
    "TWILIO_TEMPLATE_PAYMENT_RECEIVED_SID",
    "TWILIO_TEMPLATE_PAYMENT_RECORDED_SID",
    "TWILIO_TEMPLATE_PAYMENT_RECEIPT_SID"
  )
};

const razorpay = new Razorpay({
  key_id: process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET
});

function findRecordedRazorpayPayment(member, orderId, razorpayPaymentId){
  return member.payments.find(payment =>
    (orderId && payment.razorpayOrderId === orderId)
    || (razorpayPaymentId && payment.razorpayPaymentId === razorpayPaymentId)
  );
}

function applyOnlinePayment(member, selectedPaymentId, amountPaise, references = {}){
  const pendingPayments = member.payments.filter(
    payment => payment.status !== "paid" && moneyToPaise(payment.amount) > 0
  );
  const pendingTotalPaise = pendingPayments.reduce(
    (sum, payment) => sum + moneyToPaise(payment.amount),
    0
  );

  if(amountPaise <= 0){
    throw new Error("Enter a valid amount");
  }

  if(amountPaise > pendingTotalPaise){
    throw new Error("Amount cannot exceed pending due");
  }

  const selectedPayment = pendingPayments.find(
    payment => payment._id.toString() === selectedPaymentId
  );

  if(!selectedPayment){
    throw new Error("Payment not found");
  }

  const orderedPendingPayments = [
    selectedPayment,
    ...pendingPayments.filter(payment => payment._id.toString() !== selectedPaymentId)
  ];
  const paidDate = new Date();
  const appliedPayments = [];
  let remainingPaise = amountPaise;

  for(const pendingPayment of orderedPendingPayments){
    if(remainingPaise <= 0) break;

    const duePaise = moneyToPaise(pendingPayment.amount);
    if(duePaise <= 0) continue;

    const appliedPaise = Math.min(duePaise, remainingPaise);
    const appliedAmount = paiseToMoney(appliedPaise);
    const paymentReferences = {
      razorpayOrderId: references.orderId,
      razorpayPaymentId: references.paymentId
    };

    if(appliedPaise >= duePaise){
      pendingPayment.status = "paid";
      pendingPayment.paidDate = paidDate;
      pendingPayment.paymentMethod = "online";
      pendingPayment.razorpayOrderId = paymentReferences.razorpayOrderId;
      pendingPayment.razorpayPaymentId = paymentReferences.razorpayPaymentId;
      appliedPayments.push(pendingPayment);
    } else {
      pendingPayment.amount = paiseToMoney(duePaise - appliedPaise);
      member.payments.push({
        month: pendingPayment.month,
        year: pendingPayment.year,
        amount: appliedAmount,
        description: pendingPayment.description,
        chargeType: pendingPayment.chargeType,
        paymentMethod: "online",
        status: "paid",
        createdAt: pendingPayment.createdAt || paidDate,
        paidDate,
        ...paymentReferences
      });
      appliedPayments.push(member.payments[member.payments.length - 1]);
    }

    remainingPaise -= appliedPaise;
  }

  if(remainingPaise > 0){
    throw new Error("Amount cannot exceed pending due");
  }

  return appliedPayments;
}

async function getCapturedRazorpayPayment(orderId){
  const [order, paymentCollection] = await Promise.all([
    razorpay.orders.fetch(orderId),
    razorpay.orders.fetchPayments(orderId)
  ]);
  const payments = paymentCollection?.items || [];
  const capturedPayments = payments.filter(payment =>
    payment.status === "captured" && Number(payment.amount_refunded || 0) === 0
  );

  if(order.currency !== "INR"){
    throw new Error("Only INR orders can be reconciled");
  }

  if(order.status !== "paid" || capturedPayments.length !== 1){
    throw new Error("Razorpay order does not have exactly one captured payment");
  }

  const payment = capturedPayments[0];

  if(payment.order_id !== order.id || Number(payment.amount) !== Number(order.amount_paid)){
    throw new Error("Razorpay payment details do not match the order");
  }

  return { order, payment };
}

const app = express();

// Behind Render/Vercel proxies — needed for correct client IP in rate limiting
app.set("trust proxy", 1);

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cors());

// 🔒 Throttle brute-force. Generous on /login so a whole society sharing one
// WiFi/IP is never blocked, tighter on OTP/email endpoints where abuse matters.
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100,                 // plenty for a shared-IP society; still stops bots
  standardHeaders: true,
  legacyHeaders: false,
  message: { success: false, message: "Too many attempts. Please try again in a few minutes." }
});

const otpLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 20,                  // OTP/email endpoints: no real user needs many
  standardHeaders: true,
  legacyHeaders: false,
  message: { success: false, message: "Too many attempts. Please try again in a few minutes." }
});

const SECRET = process.env.JWT_SECRET;

if (!SECRET) {
  console.error("❌ FATAL: JWT_SECRET is not set. Refusing to start.");
  process.exit(1);
}

const TOKEN_TTL = process.env.JWT_EXPIRES_IN || "30d";

// 🔥 AUTO REMINDER EVERY DAY AT 10 AM
cron.schedule("0 10 * * *", async () => {
  console.log("⏰ Running auto reminder...");

  const members = await Member.find({ role: "owner" });
  const today = new Date().toDateString();

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

Open the app to pay:
${APP_PUBLIC_URL}

Please pay as soon as possible 🙏`
    ,
    {
      contentSid: TWILIO_TEMPLATES.reminder,
      variables: {
        "1": m.name,
        "2": String(profileDue),
        "3": APP_PUBLIC_URL
      }
    }
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
.then(()=>{
  console.log("✅ DB Connected");
  // Recovery path for the admin login, if ADMIN_FLAT/ADMIN_PASSWORD are set.
  seedAdminFromEnv();
  // Catch up any monthly bills the cron may have missed while asleep — but not
  // immediately. A cold start is triggered by somebody waiting on a request
  // (usually a login), and replaying bills walks every member. Serve them first.
  setTimeout(runMonthlyBillingCatchUp, 60 * 1000);
})
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

// Flat numbers are typed by hand ("a-101" vs "A-101", "admin" vs "Admin") and
// /add-member already blocks duplicates case-insensitively — so lookups must
// match the same way, or a member gets locked out of their own account.
async function findMemberByFlat(flatNumber){
  const cleanFlatNumber = String(flatNumber || "").trim();

  if(!cleanFlatNumber) return null;

  const exactMatch = await Member.findOne({ flatNumber: cleanFlatNumber });

  if(exactMatch) return exactMatch;   // indexed fast path

  return Member.findOne({
    flatNumber: { $regex: `^${escapeRegex(cleanFlatNumber)}$`, $options: "i" }
  });
}

// Nothing in the app can create or recover the admin login — /add-member always
// makes owners — so the only way back in is one the operator controls: set
// ADMIN_FLAT + ADMIN_PASSWORD in the Render dashboard and restart. Creates the
// admin if it is missing, resets its password if it already exists, and never
// touches a resident's account.
async function seedAdminFromEnv(){
  const flatNumber = String(process.env.ADMIN_FLAT || "").trim();
  const password = String(process.env.ADMIN_PASSWORD || "");

  if(!flatNumber || !password) return;

  if(password.length < 8){
    console.log("⚠️ ADMIN_PASSWORD must be at least 8 characters — admin seed skipped.");
    return;
  }

  try{
    const existing = await findMemberByFlat(flatNumber);

    // Converting an owner would hand their flat — and their payment history —
    // to whoever holds this password. Refuse and say so.
    if(existing && existing.role !== "admin"){
      console.log(`⚠️ ADMIN_FLAT "${flatNumber}" already belongs to a ${existing.role || "member"} — admin seed skipped.`);
      return;
    }

    const hashedPassword = await bcrypt.hash(password, 10);

    if(existing){
      existing.password = hashedPassword;
      await existing.save();
      console.log(`🔑 Admin password reset for flat ${existing.flatNumber}`);
      return;
    }

    // No maintenance amount, so monthly billing skips this account.
    await new Member({
      name: "Society Admin",
      flatNumber,
      password: hashedPassword,
      role: "admin",
      payments: []
    }).save();

    console.log(`🔑 Admin account created for flat ${flatNumber}`);
  }catch(err){
    console.log("Admin seed error:", err.message);
  }
}

// Same story as findMemberByFlat: emails written before normalization existed
// (or typed straight into Atlas) can carry capitals, and an exact lowercase
// match then tells an owner "no account found" for their own address.
async function findMemberByEmail(email, extraQuery = {}){
  const cleanEmail = normalizeEmail(email);

  if(!cleanEmail) return null;

  const exactMatch = await Member.findOne({ email: cleanEmail, ...extraQuery });

  if(exactMatch) return exactMatch;   // indexed fast path

  return Member.findOne({
    email: { $regex: `^${escapeRegex(cleanEmail)}$`, $options: "i" },
    ...extraQuery
  });
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

function getMailFrom(){
  return process.env.MAIL_FROM || process.env.EMAIL_USER;
}

// Render blocks outbound SMTP, which is the whole reason mail was being bounced
// through a relay on Vercel. Resend delivers over ordinary HTTPS, so the
// backend can send directly and the relay hop disappears.
// Pasting a key into a dashboard picks up trailing spaces and newlines far too
// easily, and the only symptom is a bare 401.
function getResendApiKey(){
  return String(process.env.RESEND_API_KEY || "").trim();
}

function shouldUseResend(){
  return Boolean(getResendApiKey() && getMailFrom());
}

async function sendViaResend(mailOptions){
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${getResendApiKey()}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      from: getMailFrom(),
      to: [mailOptions.to],
      subject: mailOptions.subject,
      text: mailOptions.text
    }),
    signal: AbortSignal.timeout(15000)
  });

  if(!response.ok){
    const message = await response.text();
    throw new Error(`Resend rejected the email (${response.status}): ${message}`);
  }
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

// Resend first when it is configured, then the relay, then direct SMTP for
// local development. One place decides, so both send paths agree.
function getMailTransport(){
  if(shouldUseResend()) return "resend";
  if(shouldUseMailRelay()) return "relay";
  return "smtp";
}

// Try every configured transport in order rather than giving up on the first
// failure — a resident who never receives their OTP is locked out of the app
// entirely, so one flaky provider must not decide that.
async function deliverMail(mailOptions){
  const transports = [];

  if(shouldUseResend()){
    transports.push(["resend", () => sendViaResend(mailOptions)]);
  }

  if(shouldUseMailRelay()){
    transports.push(["relay", () => relayMail(mailOptions)]);
  }

  transports.push(["smtp", async () => {
    const transporter = await createMailTransporter();
    await transporter.sendMail(mailOptions);
  }]);

  let lastError;

  for(const [name, send] of transports){
    try{
      await send();
      return name;
    }catch(error){
      lastError = error;
      console.log(`Mail via ${name} failed:`, error.message);
    }
  }

  throw lastError || new Error("No mail transport is configured");
}

function sendMailInBackground(mailOptions, context){
  deliverMail(mailOptions)
    .then(via => console.log(`${context} email sent via ${via}`))
    .catch(error => console.log(`${context} email failed:`, error.message));
}

async function sendMailNow(mailOptions, context){
  const transport = getMailTransport();

  try {
    const via = await deliverMail(mailOptions);

    console.log(`${context} email sent via ${via}`);
  } catch (error) {
    console.log(`${context} email failed:`, error.message);

    throw new Error(
      transport === "resend"
        ? "Email could not be sent. Check RESEND_API_KEY and MAIL_FROM in Render."
        : "Email could not be sent. Check EMAIL_USER and EMAIL_PASS in Render."
    );
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
app.post("/login", loginLimiter, async(req,res)=>{
 try{
  const flatNumber = String(req.body.flatNumber || "").trim();
  const password = String(req.body.password || "");

  if(!flatNumber || !password) return res.json({success:false});

  const user=await findMemberByFlat(flatNumber);
  if(!user || !user.password) return res.json({success:false});

  const match=await bcrypt.compare(password,user.password);
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

  // New residents are handed flat number as both username and password, so a
  // password still equal to the flat number means they have never set one.
  // No extra field to keep in sync: it stops being true the moment they do.
  const usingDefaultPassword = user.role === "owner"
    ? await bcrypt.compare(user.flatNumber, user.password)
    : false;

  const token=jwt.sign({flat:user.flatNumber,role:user.role},SECRET,{expiresIn:TOKEN_TTL});
  res.json({success:true,token,role:user.role,mustChangePassword:usingDefaultPassword});
 }catch(err){
  // A cold-starting free-tier instance can reject the very first DB query.
  // Say that plainly instead of dying with an unreadable 500.
  console.log("Login error:", err.message);
  res.status(503).json({
    success:false,
    message:"Server is still waking up. Please try again in a few seconds."
  });
 }
});

app.post("/verify-email", otpLimiter, async(req,res)=>{
  const { flatNumber, otp } = req.body;

  const user = await findMemberByFlat(flatNumber);

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

  const reminderResult = await sendWhatsApp(
    formatPhone(member.phone),
    `❗ Reminder

Hello ${member.name},

You have pending maintenance of ₹${profileDue}.

Open the app to pay:
${APP_PUBLIC_URL}

Please pay as soon as possible 🙏`
    ,
    {
      contentSid: TWILIO_TEMPLATES.reminder,
      variables: {
        "1": member.name,
        "2": String(profileDue),
        "3": APP_PUBLIC_URL
      }
    }
  );

  if(!reminderResult.ok){
    return res.status(500).send(`Reminder failed: ${reminderResult.error || reminderResult.code || "Twilio rejected the message"}`);
  }

  res.send(`Reminder queued: ${reminderResult.sid}`);
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
  const { phone, email } = req.body;

  // 🔒 Always update the logged-in user's own record, never an arbitrary id
  const user = await Member.findOne({ flatNumber: req.user.flat });

  if (!user) {
    return res.status(404).send("Member not found");
  }

  const cleanEmail = normalizeEmail(email);

  if (cleanEmail && !/^\S+@\S+\.\S+$/.test(cleanEmail)) {
    return res.status(400).send("Please enter a valid email address");
  }

  if (phone !== undefined) {
    user.phone = String(phone).trim();
  }

  // Changing email must re-trigger verification for owners
  if (cleanEmail && cleanEmail !== normalizeEmail(user.email)) {
    user.email = cleanEmail;

    if (user.role === "owner") {
      user.emailVerified = false;
      user.emailVerificationOtp = null;
      user.emailVerificationExpiry = null;
    }
  }

  await user.save();

  res.send("Profile updated");
});

// Read-only view of the logged-in account. Admins have no resident dashboard,
// so this is how they see their own details.
app.get("/me", auth, async (req, res) => {
  const user = await Member.findOne({ flatNumber: req.user.flat })
    .select("name flatNumber phone email role emailVerified");

  if (!user) return res.status(404).send("Member not found");

  res.json(user);
});

// Change your own password from inside the app. Until now the only way to
// change one was the emailed reset link, which the admin account cannot always
// use. Scoped to the logged-in user — no id is accepted from the body — and it
// writes nothing but the password field.
app.post("/change-password", auth, otpLimiter, async (req, res) => {
  try {
    const currentPassword = String(req.body.currentPassword || "");
    const newPassword = String(req.body.newPassword || "");

    if (!currentPassword || !newPassword) {
      return res.status(400).send("Enter your current and new password");
    }

    if (newPassword.length < 8) {
      return res.status(400).send("New password must be at least 8 characters");
    }

    const user = await Member.findOne({ flatNumber: req.user.flat });

    if (!user || !user.password) return res.status(404).send("Member not found");

    const match = await bcrypt.compare(currentPassword, user.password);

    if (!match) return res.status(400).send("Current password is wrong");

    user.password = await bcrypt.hash(newPassword, 10);

    await user.save();

    console.log(`🔑 Password changed by ${user.flatNumber}`);

    res.send("Password updated");
  } catch (error) {
    console.log("Change password error:", error.message);
    res.status(500).send("Could not update password");
  }
});

async function sendPasswordResetLink(req, res){
  try {
    const { email } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail) return res.status(400).send("Enter email");

    const user = await findMemberByEmail(cleanEmail);

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

app.post("/send-reset-link", otpLimiter, sendPasswordResetLink);

app.post("/send-otp", otpLimiter, sendPasswordResetLink);

app.post("/reset-password", otpLimiter, async (req, res) => {
  try {
    const { email, token, newPassword } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail || !token || !newPassword) {
      return res.status(400).send("Open the reset link and enter a new password");
    }

    if (String(newPassword).length < 4) {
      return res.status(400).send("Password must be at least 4 characters");
    }

    const user = await findMemberByEmail(cleanEmail, {
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

app.post("/verify-otp", otpLimiter, async (req, res) => {
  try {
    const { email, otp, newPassword } = req.body;
    const cleanEmail = normalizeEmail(email);

    if (!cleanEmail || !otp || !newPassword) {
      return res.status(400).send("Enter email, OTP, and new password");
    }

    const user = await findMemberByEmail(cleanEmail);

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
  try {
    const { memberId, month, year, amount, description, chargeType } = req.body;
    const cleanAmount = Number(amount);
    const isCustomCharge = chargeType === "custom";
    const cleanDescription = String(description || "").trim().slice(0, 160);
    const cleanMonth = isCustomCharge ? "Custom charge" : String(month || "").trim();
    const cleanYear = Number(year || new Date().getFullYear());

    if (!memberId || !Number.isFinite(cleanAmount) || cleanAmount <= 0) {
      return res.status(400).send("Enter a valid charge amount");
    }

    if (!isCustomCharge && (!cleanMonth || !cleanYear)) {
      return res.status(400).send("Select a valid due range");
    }

    if (isCustomCharge && !cleanDescription) {
      return res.status(400).send("Enter charge remarks");
    }

    const m = await Member.findById(memberId);

    if (!m) {
      return res.status(404).send("Member not found");
    }

    const label = isCustomCharge
      ? cleanDescription
      : `Maintenance charge for ${cleanMonth} ${cleanYear}`;

    m.payments.push({
      month: cleanMonth,
      year: cleanYear,
      amount: cleanAmount,
      description: label,
      chargeType: isCustomCharge ? "custom" : "maintenance",
      status: "pending",
      createdAt: new Date()
    });

    await m.save();

    await sendWhatsApp(
      formatPhone(m.phone),
      `❗ New Due Added

📅 ${label}
💰 Amount: ₹${cleanAmount}

Open the app to pay:
${APP_PUBLIC_URL}

Please pay on time 🙏`
      ,
      {
        contentSid: TWILIO_TEMPLATES.dueAdded,
        variables: {
          "1": m.name,
          "2": label,
          "3": String(cleanAmount),
          "4": APP_PUBLIC_URL
        }
      }
    );

    res.send("Due added");
  } catch (error) {
    console.log("Add due error:", error);
    res.status(500).send("Charge could not be added");
  }
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
  const members = await Member.find({ role: "owner" });

  let list = [];

  members.forEach(m => {
    (m.payments || []).forEach(p => {
      list.push({
        name: m.name,
        flat: m.flatNumber,
        month: p.month,
        year: p.year,
        amount: p.amount,
        description: p.description,
        chargeType: p.chargeType,
        paymentMethod: p.paymentMethod,
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
      { header: "Description", key: "description", width: 32 },
      { header: "Payment Method", key: "paymentMethod", width: 18 },
      { header: "Amount", key: "amount", width: 15 },
      { header: "Status", key: "status", width: 15 }
    ];

    data.forEach(p => {
      ws.addRow({
        name: p.name,
        flat: p.flat,
        month: p.month,
        year: p.year,
        description: p.description || "",
        paymentMethod: p.paymentMethod || "",
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

  if(!user) return res.status(404).send("Member not found");

  const now = new Date();
  const month = now.toLocaleString("default",{month:"long"});
  const year = now.getFullYear();

  // 🔥 CHECK IF CURRENT MONTH BILL EXISTS
  // Only residents get billed, and only when a real maintenance amount is set.
  // Without these guards this quietly turned a staff/admin login into a paying
  // customer, and pushed an amount-less due onto anyone with no amount.
  const monthlyAmount = Number(user.monthlyMaintenance);
  const billable = user.role === "owner" && Number.isFinite(monthlyAmount) && monthlyAmount > 0;

  if(billable){
    let exists = user.payments.find(p=>p.month===month && p.year===year);

    if(!exists){
      user.payments.push({
        month,
        year,
        amount: monthlyAmount
      });

      await user.save();
    }
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

app.post("/record-cash-payment", auth, adminOnly, async(req,res)=>{
  try {
    const { paymentId, amount } = req.body;

    if (!paymentId) {
      return res.status(400).send("Payment is required");
    }

    const member = await Member.findOne({ "payments._id": paymentId });

    if (!member) {
      return res.status(404).send("Payment not found");
    }

    const payment = member.payments.find(p => p._id.toString() === paymentId);

    if (!payment) {
      return res.status(404).send("Payment not found");
    }

    if (payment.status === "paid") {
      return res.status(400).send("This due is already paid");
    }

    const duePaise = moneyToPaise(payment.amount);
    const amountPaise = amount === undefined || amount === null || amount === ""
      ? duePaise
      : moneyToPaise(amount);

    if (duePaise <= 0) {
      return res.status(400).send("Due amount is invalid");
    }

    if (amountPaise <= 0) {
      return res.status(400).send("Cash amount must be more than zero");
    }

    if (amountPaise > duePaise) {
      return res.status(400).send("Cash amount cannot exceed this due");
    }

    const paidDate = new Date();
    const paidAmount = paiseToMoney(amountPaise);
    let receiptPayment = payment;

    if (amountPaise === duePaise) {
      payment.status = "paid";
      payment.paidDate = paidDate;
      payment.paymentMethod = "cash";
    } else {
      payment.amount = paiseToMoney(duePaise - amountPaise);
      member.payments.push({
        month: payment.month,
        year: payment.year,
        amount: paidAmount,
        description: payment.description,
        chargeType: payment.chargeType,
        paymentMethod: "cash",
        status: "paid",
        createdAt: payment.createdAt || paidDate,
        paidDate
      });
      receiptPayment = member.payments[member.payments.length - 1];
    }

    await member.save();

    const remainingDue = member.payments.reduce(
      (sum, p) => p.status !== "paid" ? sum + Number(p.amount || 0) : sum,
      0
    );
    const paymentLabel = receiptPayment.description || `${receiptPayment.month} ${receiptPayment.year}`;
    const billFile = await generateBill(member, receiptPayment);
    const billUrl = getBillUrl(billFile);

    const whatsappResult = await sendWhatsApp(
      formatPhone(member.phone),
      `✅ Cash Payment Received!

📅 Charge: ${paymentLabel}
💰 Amount: ₹${paidAmount}
⏳ Remaining Due: ₹${remainingDue}

${billUrl ? `📄 Download Bill:
${billUrl}
` : ""}
Open the app:
${APP_PUBLIC_URL}

Thank you 🙏`,
      {
        contentSid: TWILIO_TEMPLATES.paymentReceived,
        variables: {
          "1": member.name,
          "2": String(paidAmount),
          "3": String(remainingDue),
          "4": APP_PUBLIC_URL
        }
      }
    );

    if(!whatsappResult.ok){
      console.log("Cash payment WhatsApp failed:", whatsappResult);
    }

    res.send(amountPaise === duePaise
      ? "Cash payment recorded"
      : `Cash payment recorded. Remaining due: ₹${remainingDue}`);
  } catch (error) {
    console.log("Cash payment error:", error);
    res.status(500).send("Cash payment could not be recorded");
  }
});

app.delete("/due/:paymentId", auth, adminOnly, async(req,res)=>{
  try {
    const { paymentId } = req.params;

    if (!paymentId) {
      return res.status(400).send("Payment is required");
    }

    const member = await Member.findOne({ "payments._id": paymentId });

    if (!member) {
      return res.status(404).send("Due not found");
    }

    const payment = member.payments.find(p => p._id.toString() === paymentId);

    if (!payment) {
      return res.status(404).send("Due not found");
    }

    if (payment.status === "paid") {
      return res.status(400).send("Paid receipts cannot be deleted");
    }

    member.payments.pull(payment._id);
    await member.save();

    res.send("Due deleted");
  } catch (error) {
    console.log("Delete due error:", error);
    res.status(500).send("Due could not be deleted");
  }
});


// ================= PENDING =================
app.get("/pending", auth, adminOnly, async(req,res)=>{
  const members=await Member.find({ role: "owner" });

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
    members = await Member.find({ role: "owner" });   // 👈 FILTER ONLY OWNERS
  }

  console.log("MEMBERS COUNT:", members.length);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Report");

  ws.columns = [
    { header: "Name", key: "name" },
    { header: "Flat", key: "flat" },
    { header: "Month", key: "month" },
    { header: "Year", key: "year" },
    { header: "Description", key: "description" },
    { header: "Payment Method", key: "paymentMethod" },
    { header: "Amount", key: "amount" },
    { header: "Status", key: "status" }
  ];

  members.forEach(m => {
    m.payments.forEach(p => {
      const monthNumber = monthToNumber(p.month);

      const pDate = Number.isFinite(monthNumber)
        ? new Date(p.year, monthNumber - 1, 1)
        : new Date(p.createdAt || Date.now());

      if (pDate >= start && pDate <= end) {
        ws.addRow({
          name: m.name,
          flat: m.flatNumber,
          month: p.month,
          year: p.year,
          description: p.description || "",
          paymentMethod: p.paymentMethod || "",
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

// 🔥 AUTO MONTHLY BILL
// Idempotent: adds the current month's maintenance to every member who does
// not already have it. Safe to run any number of times — a member is never
// billed twice for the same month. Each member is isolated in its own
// try/catch, so one failure can never stop the rest of the batch (that was
// the old bug that billed "some members but not all").
async function ensureMonthlyBillsForAll({ notify = false } = {}) {
  const now = new Date();
  const currentMonth = now.toLocaleString("default", { month: "long" });
  const currentYear = now.getFullYear();

  const members = await Member.find();
  const report = {
    month: `${currentMonth} ${currentYear}`,
    billed: [],
    skipped: [],
    failed: []
  };

  for (const m of members) {
    try {
      const amount = Number(m.monthlyMaintenance);

      // Only residents owe maintenance. Staff/admin logins are not customers,
      // even if an amount was left on their record.
      if (m.role !== "owner") {
        report.skipped.push({ flat: m.flatNumber, reason: "not a resident" });
        continue;
      }

      // Skip anyone without a real maintenance amount
      if (!Number.isFinite(amount) || amount <= 0) {
        report.skipped.push({ flat: m.flatNumber, reason: "no maintenance amount" });
        continue;
      }

      const alreadyExists = (m.payments || []).some(
        p => p.month === currentMonth && p.year === currentYear
      );

      if (alreadyExists) {
        report.skipped.push({ flat: m.flatNumber, reason: "already billed" });
        continue;
      }

      m.payments.push({
        month: currentMonth,
        year: currentYear,
        amount,
        status: "pending",
        createdAt: new Date()
      });

      await m.save();
      report.billed.push({ flat: m.flatNumber, name: m.name, amount });
      console.log(`✅ Bill added for ${m.name} (${m.flatNumber})`);

      if (notify) {
        // Fire-and-forget: a slow or failing WhatsApp must never block or
        // skip the billing of the next member.
        sendWhatsApp(
          formatPhone(m.phone),
          `📅 Monthly Maintenance Added

💰 ₹${amount}

Open the app to pay:
${APP_PUBLIC_URL}

Please pay within 5 days 🙏`,
          {
            contentSid: TWILIO_TEMPLATES.monthlyDue,
            variables: {
              "1": m.name,
              "2": `${currentMonth} ${currentYear}`,
              "3": String(amount),
              "4": APP_PUBLIC_URL
            }
          }
        ).catch(err =>
          console.log(`Monthly bill WhatsApp failed for ${m.flatNumber}:`, err.message)
        );
      }
    } catch (err) {
      report.failed.push({ flat: m.flatNumber, error: err.message });
      console.log(`❌ Monthly bill failed for ${m.flatNumber}:`, err.message);
    }
  }

  console.log(
    `🎉 Monthly billing (${report.month}): ${report.billed.length} billed, ` +
    `${report.skipped.length} skipped, ${report.failed.length} failed`
  );

  return report;
}

// Runs on the 1st of every month (when the server happens to be awake).
cron.schedule("0 0 1 * *", () => {
  console.log("📅 Monthly billing cron triggered...");
  ensureMonthlyBillsForAll({ notify: true }).catch(err =>
    console.log("❌ Monthly billing cron error:", err.message)
  );
});

// 🩹 Self-heal: on every startup, catch up any current-month bills that the
// cron missed (e.g. the free-tier server was asleep at midnight on the 1st).
// Silent by design — the existing daily 10am reminder cron notifies members
// of any pending dues, so this avoids a WhatsApp burst on every cold start.
async function runMonthlyBillingCatchUp() {
  try {
    console.log("🩹 Monthly billing catch-up on startup...");
    await ensureMonthlyBillsForAll({ notify: false });
  } catch (err) {
    console.log("❌ Monthly billing catch-up error:", err.message);
  }
}

// Admin can trigger billing on demand and see exactly who was billed/skipped.
app.post("/admin/run-monthly-billing", auth, adminOnly, async (req, res) => {
  try {
    const notify = req.body?.notify !== false; // notify unless explicitly false
    const report = await ensureMonthlyBillsForAll({ notify });
    res.json({ ok: true, ...report });
  } catch (err) {
    console.log("Manual monthly billing failed:", err.message);
    res.status(500).json({ ok: false, error: "Monthly billing failed" });
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
    receipt: "receipt_" + Date.now(),
    notes: {
      memberId: member._id.toString(),
      flatNumber: member.flatNumber,
      paymentId: payment._id.toString()
    }
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

app.post("/verify-payment", auth, async (req, res) => {
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
      flatNumber: req.user.flat,
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

    let capturedOrder;
    let capturedPayment;
    try {
      const captured = await getCapturedRazorpayPayment(razorpay_order_id);
      capturedOrder = captured.order;
      capturedPayment = captured.payment;
    } catch (error) {
      console.log("❌ Razorpay capture verification failed:", error.message);
      return res.status(400).json({ error: "Payment has not been captured by Razorpay" });
    }

    if (capturedPayment.id !== razorpay_payment_id) {
      return res.status(400).json({ error: "Razorpay payment does not match the order" });
    }

    const amountPaise = Number(capturedPayment.amount);
    const requestedAmountPaise = moneyToPaise(paidAmount ?? payment.amount);

    if (amountPaise <= 0 || amountPaise !== requestedAmountPaise) {
      console.log("❌ Amount mismatch:", amountPaise, requestedAmountPaise);
      return res.status(400).json({ error: "Payment amount mismatch" });
    }

    const orderNotes = capturedOrder.notes || {};
    if(
      (orderNotes.memberId && String(orderNotes.memberId) !== member._id.toString())
      || (orderNotes.flatNumber && String(orderNotes.flatNumber) !== member.flatNumber)
      || (orderNotes.paymentId && String(orderNotes.paymentId) !== paymentId)
    ){
      return res.status(400).json({ error: "Payment order belongs to a different due" });
    }

    const recordedPayment = findRecordedRazorpayPayment(
      member,
      razorpay_order_id,
      razorpay_payment_id
    );

    if(recordedPayment){
      return res.json({
        success: true,
        alreadyRecorded: true,
        paidAmount: Number(recordedPayment.amount || 0)
      });
    }

    const appliedPayments = applyOnlinePayment(member, paymentId, amountPaise, {
      orderId: razorpay_order_id,
      paymentId: razorpay_payment_id
    });

    await member.save();

    console.log("✅ DB UPDATED");

    const paidTotal = paiseToMoney(amountPaise);
    const receiptPayment = appliedPayments.length === 1
      ? appliedPayments[0]
      : {
          _id: razorpay_payment_id,
          month: "Multiple dues",
          year: new Date().getFullYear(),
          amount: paidTotal,
          description: "Multiple dues payment"
        };
    const remainingDue = member.payments.reduce(
      (sum, p) => p.status !== "paid" ? sum + Number(p.amount || 0) : sum,
      0
    );

    // 📄 generate bill
    const billFile = await generateBill(member, receiptPayment);
    const billUrl = getBillUrl(billFile);

    const paymentLabel = receiptPayment.description
      || `${receiptPayment.month} ${receiptPayment.year}`;

    // 📲 send whatsapp
    const whatsappResult = await sendWhatsApp(
      formatPhone(member.phone),
      `✅ Payment Received!

📅 Charge: ${paymentLabel}
💰 Amount: ₹${paidTotal}
⏳ Remaining Due: ₹${remainingDue}

${billUrl ? `📄 Download Bill:
${billUrl}
` : ""}
Open the app:
${APP_PUBLIC_URL}

Thank you 🙏`
      ,
      {
        contentSid: TWILIO_TEMPLATES.paymentReceived,
        variables: {
          "1": member.name,
          "2": String(paidTotal),
          "3": String(remainingDue),
          "4": APP_PUBLIC_URL
        }
      }
    );

    if(whatsappResult.ok){
      console.log("✅ WhatsApp sent", whatsappResult.sid);
    } else {
      console.log("❌ WhatsApp payment receipt failed:", whatsappResult.error || whatsappResult.code);
    }

    res.json({ success: true, paidAmount: paidTotal, remainingDue, whatsapp: whatsappResult });

  } catch (err) {
    console.log("❌ ERROR:", err);
    res.status(500).json({ error: "Verification failed" });
  }
});

app.get("/admin/razorpay-order/:orderId", auth, adminOnly, async(req, res) => {
  try {
    const { orderId } = req.params;

    if(!/^order_[A-Za-z0-9]+$/.test(orderId || "")){
      return res.status(400).json({ error: "Invalid Razorpay order ID" });
    }

    const order = await razorpay.orders.fetch(orderId);
    const paymentCollection = await razorpay.orders.fetchPayments(orderId);
    const payments = (paymentCollection?.items || []).map(payment => ({
      id: payment.id,
      amount: paiseToMoney(payment.amount),
      amountRefunded: paiseToMoney(payment.amount_refunded),
      status: payment.status,
      captured: payment.captured,
      orderId: payment.order_id,
      createdAt: payment.created_at
        ? new Date(payment.created_at * 1000).toISOString()
        : null
    }));

    res.json({
      id: order.id,
      amount: paiseToMoney(order.amount),
      amountPaid: paiseToMoney(order.amount_paid),
      amountDue: paiseToMoney(order.amount_due),
      currency: order.currency,
      status: order.status,
      notes: order.notes || {},
      payments
    });
  } catch(error) {
    console.log("Razorpay order diagnostic failed:", error.message);
    res.status(400).json({ error: "Razorpay order could not be fetched" });
  }
});

app.post("/admin/reconcile-razorpay-payment", auth, adminOnly, async(req, res) => {
  try {
    const { orderId, memberId, paymentId } = req.body;

    if(!orderId || !memberId || !paymentId){
      return res.status(400).json({
        error: "Order, resident, and due payment are required"
      });
    }

    const member = await Member.findOne({
      _id: memberId,
      "payments._id": paymentId
    });

    if(!member){
      return res.status(404).json({ error: "Resident payment not found" });
    }

    const { order, payment } = await getCapturedRazorpayPayment(orderId);
    const existingMember = await Member.findOne({
      $or: [
        { "payments.razorpayOrderId": order.id },
        { "payments.razorpayPaymentId": payment.id }
      ]
    });

    if(existingMember && existingMember._id.toString() !== member._id.toString()){
      return res.status(409).json({
        error: "This Razorpay payment is already recorded for another resident"
      });
    }

    const recordedPayment = findRecordedRazorpayPayment(member, order.id, payment.id);

    if(recordedPayment){
      const remainingDue = member.payments.reduce(
        (sum, item) => item.status !== "paid" ? sum + Number(item.amount || 0) : sum,
        0
      );

      return res.json({
        success: true,
        alreadyRecorded: true,
        paidAmount: Number(recordedPayment.amount || 0),
        remainingDue
      });
    }

    const amountPaise = Number(payment.amount);
    const appliedPayments = applyOnlinePayment(member, paymentId, amountPaise, {
      orderId: order.id,
      paymentId: payment.id
    });

    await member.save();

    const remainingDue = member.payments.reduce(
      (sum, item) => item.status !== "paid" ? sum + Number(item.amount || 0) : sum,
      0
    );
    const paidTotal = paiseToMoney(amountPaise);
    const receiptPayment = appliedPayments.length === 1
      ? appliedPayments[0]
      : {
          _id: payment.id,
          month: "Multiple dues",
          year: new Date().getFullYear(),
          amount: paidTotal,
          description: "Recovered Razorpay payment"
        };

    let whatsappResult = { ok: false, error: "Receipt notification was not attempted" };
    try {
      const billFile = await generateBill(member, receiptPayment);
      const billUrl = getBillUrl(billFile);
      const paymentLabel = receiptPayment.description
        || `${receiptPayment.month} ${receiptPayment.year}`;

      whatsappResult = await sendWhatsApp(
        formatPhone(member.phone),
        `✅ Payment Received!

📅 Charge: ${paymentLabel}
💰 Amount: ₹${paidTotal}
⏳ Remaining Due: ₹${remainingDue}

${billUrl ? `📄 Download Bill:
${billUrl}
` : ""}
Open the app:
${APP_PUBLIC_URL}

Thank you 🙏`,
        {
          contentSid: TWILIO_TEMPLATES.paymentReceived,
          variables: {
            "1": member.name,
            "2": String(paidTotal),
            "3": String(remainingDue),
            "4": APP_PUBLIC_URL
          }
        }
      );
    } catch(notificationError) {
      console.log("Recovered payment receipt failed:", notificationError.message);
      whatsappResult = { ok: false, error: notificationError.message };
    }

    res.json({
      success: true,
      paidAmount: paidTotal,
      remainingDue,
      razorpayPaymentId: payment.id,
      whatsapp: whatsappResult,
      appliedPayments: appliedPayments.map(item => ({
        id: item._id,
        amount: item.amount,
        month: item.month,
        year: item.year,
        description: item.description
      }))
    });
  } catch(error) {
    console.log("Razorpay reconciliation failed:", error.message);
    res.status(400).json({ error: error.message || "Payment could not be reconciled" });
  }
});

async function sendWhatsApp(to, message, options = {}) {
  try {
    if(!client || !TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN){
      return {
        ok: false,
        error: "Twilio Account SID/Auth Token is not configured. Use the Account SID starting with AC, not the User SID starting with US."
      };
    }

    if(!TWILIO_ACCOUNT_SID.startsWith("AC")){
      return {
        ok: false,
        error: "Twilio Account SID looks wrong. It must start with AC."
      };
    }

    const cleanTo = normalizeWhatsAppAddress(to);
    if(!cleanTo){
      return { ok: false, error: "Recipient WhatsApp phone number is missing" };
    }

    if(!TWILIO_MESSAGING_SERVICE_SID && !WHATSAPP_FROM){
      return { ok: false, error: "Twilio WhatsApp sender is not configured" };
    }

    if(TWILIO_MESSAGING_SERVICE_SID && !TWILIO_MESSAGING_SERVICE_SID.startsWith("MG")){
      return {
        ok: false,
        error: "Twilio Messaging Service SID looks wrong. It must start with MG, or remove it and use TWILIO_WHATSAPP_FROM."
      };
    }

    const contentSid = String(options.contentSid || "").trim();
    const payload = {
      to: cleanTo
    };

    if(TWILIO_MESSAGING_SERVICE_SID){
      payload.messagingServiceSid = TWILIO_MESSAGING_SERVICE_SID;
    } else {
      payload.from = WHATSAPP_FROM;
    }

    const statusCallback = getTwilioStatusCallbackUrl();
    if(statusCallback){
      payload.statusCallback = statusCallback;
    }

    if(contentSid){
      payload.contentSid = contentSid;
      payload.contentVariables = JSON.stringify(options.variables || {});
    } else {
      payload.body = message;
    }

    console.log("📲 Twilio send request:", {
      to: maskPhone(payload.to),
      from: maskPhone(payload.from),
      messagingServiceSid: maskSid(payload.messagingServiceSid),
      contentSid: maskSid(payload.contentSid),
      mode: payload.contentSid ? "template" : "body",
      statusCallback: Boolean(payload.statusCallback)
    });

    const msg = await client.messages.create(payload);

    console.log("✅ Twilio message queued:", msg.sid, msg.status);
    return {
      ok: true,
      sid: msg.sid,
      status: msg.status,
      to: maskPhone(payload.to),
      from: maskPhone(payload.from),
      messagingServiceSid: maskSid(payload.messagingServiceSid),
      contentSid: maskSid(payload.contentSid)
    };
  } catch (err) {
    console.log("❌ Twilio send error:", {
      code: err.code,
      status: err.status,
      message: err.message
    });
    return { ok: false, error: err.message, code: err.code };
  }
}

app.post("/twilio-status", (req, res) => {
  const status = req.body.MessageStatus || req.body.SmsStatus;

  console.log("📬 Twilio delivery status:", {
    sid: maskSid(req.body.MessageSid || req.body.SmsSid),
    status,
    errorCode: req.body.ErrorCode,
    errorMessage: req.body.ErrorMessage,
    to: maskPhone(req.body.To),
    from: maskPhone(req.body.From)
  });

  res.sendStatus(204);
});

app.get("/test-whatsapp", async (req, res) => {
  if(!verifyTestSecret(req, res)) return;

  const template = String(req.query.template || "").trim();
  const to = formatPhone(req.query.to || process.env.TWILIO_TEST_TO || "+918670433655");
  const templateOptions = getTestTemplateOptions(template);

  if(template && !templateOptions){
    return res.status(400).json({
      ok: false,
      error: "Unknown or unconfigured template",
      configuredTemplates: getTwilioTemplateDiagnostics()
    });
  }

  const result = await sendWhatsApp(
    to,
    "WhatsApp working from New Town Society.",
    templateOptions || {}
  );

  if(!result.ok){
    return res.status(500).json(result);
  }

  res.json(result);
});

app.get("/twilio-diagnostics", (req, res) => {
  if(!verifyTestSecret(req, res)) return;

  res.json({
    ok: true,
    accountSidConfigured: Boolean(TWILIO_ACCOUNT_SID),
    accountSidLooksValid: TWILIO_ACCOUNT_SID.startsWith("AC"),
    authTokenConfigured: Boolean(TWILIO_AUTH_TOKEN),
    whatsappFromConfigured: Boolean(process.env.TWILIO_WHATSAPP_FROM),
    usingSandboxSender: WHATSAPP_FROM === TWILIO_SANDBOX_FROM,
    whatsappFrom: maskPhone(WHATSAPP_FROM),
    messagingServiceConfigured: Boolean(TWILIO_MESSAGING_SERVICE_SID),
    messagingServiceLooksValid: !TWILIO_MESSAGING_SERVICE_SID || TWILIO_MESSAGING_SERVICE_SID.startsWith("MG"),
    messagingServiceSid: maskSid(TWILIO_MESSAGING_SERVICE_SID),
    appPublicUrl: APP_PUBLIC_URL,
    backendPublicUrl: BACKEND_PUBLIC_URL,
    statusCallbackUrl: getTwilioStatusCallbackUrl(),
    templates: getTwilioTemplateDiagnostics()
  });
});

app.get("/test-whatsapp-status", async (req, res) => {
  try {
    if(!verifyTestSecret(req, res)) return;

    const sid = String(req.query.sid || "").trim();
    if(!sid) return res.status(400).json({ ok: false, error: "Message SID is required" });
    if(!client) return res.status(500).json({ ok: false, error: "Twilio is not configured" });

    const message = await client.messages(sid).fetch();

    res.json({
      ok: true,
      sid: message.sid,
      status: message.status,
      errorCode: message.errorCode,
      errorMessage: message.errorMessage,
      direction: message.direction,
      from: message.from,
      to: message.to,
      dateCreated: message.dateCreated,
      dateSent: message.dateSent
    });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, code: err.code });
  }
});

app.get("/twilio-live-diagnostics", async (req, res) => {
  try {
    if(!verifyTestSecret(req, res)) return;

    const sid = String(req.query.sid || "").trim();
    res.json(await buildTwilioLiveDiagnostics(sid));
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, code: err.code });
  }
});

app.get("/admin/twilio-live-diagnostics", auth, adminOnly, async (req, res) => {
  try {
    const sid = String(req.query.sid || "").trim();
    res.json(await buildTwilioLiveDiagnostics(sid));
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, code: err.code });
  }
});

async function buildTwilioLiveDiagnostics(sid = "") {
  if(!client) return { ok: false, error: "Twilio is not configured" };

  const result = {
    ok: true,
    config: {
      accountSidConfigured: Boolean(TWILIO_ACCOUNT_SID),
      accountSidLooksValid: TWILIO_ACCOUNT_SID.startsWith("AC"),
      authTokenConfigured: Boolean(TWILIO_AUTH_TOKEN),
      whatsappFromConfigured: Boolean(process.env.TWILIO_WHATSAPP_FROM),
      whatsappFrom: maskPhone(WHATSAPP_FROM),
      usingSandboxSender: WHATSAPP_FROM === TWILIO_SANDBOX_FROM,
      messagingServiceConfigured: Boolean(TWILIO_MESSAGING_SERVICE_SID),
      messagingServiceLooksValid: !TWILIO_MESSAGING_SERVICE_SID || TWILIO_MESSAGING_SERVICE_SID.startsWith("MG"),
      messagingServiceSid: maskSid(TWILIO_MESSAGING_SERVICE_SID),
      statusCallbackUrl: getTwilioStatusCallbackUrl(),
      templates: getTwilioTemplateDiagnostics()
    },
    account: null,
    requestedMessage: null,
    recentMessages: [],
    whatsappSenders: [],
    contentTemplates: [],
    alerts: [],
    errors: {}
  };

  try {
    const account = await client.api.accounts(TWILIO_ACCOUNT_SID).fetch();
    result.account = {
      sid: maskSid(account.sid),
      status: account.status,
      type: account.type,
      friendlyName: account.friendlyName
    };
  } catch (err) {
    result.errors.account = maskTwilioError(err);
  }

  if(sid){
    try {
      const message = await client.messages(sid).fetch();
      result.requestedMessage = formatTwilioMessage(message);
    } catch (err) {
      result.errors.requestedMessage = maskTwilioError(err);
    }
  }

  try {
    const messages = await client.messages.list({ limit: 10 });
    result.recentMessages = messages.map(formatTwilioMessage);
  } catch (err) {
    result.errors.recentMessages = maskTwilioError(err);
  }

  try {
    const senders = await client.messaging.v2.channelsSenders.list({
      channel: "whatsapp",
      limit: 20
    });
    result.whatsappSenders = senders.map(formatTwilioSender);
  } catch (err) {
    result.errors.whatsappSenders = maskTwilioError(err);
  }

  try {
    const configuredTemplateSids = new Set(Object.values(TWILIO_TEMPLATES).filter(Boolean));
    const contents = await client.content.v1.contentAndApprovals.list({ limit: 100 });
    result.contentTemplates = contents
      .filter(content => configuredTemplateSids.has(content.sid))
      .map(formatTwilioContentApproval);
  } catch (err) {
    result.errors.contentTemplates = maskTwilioError(err);
  }

  try {
    const alerts = await client.monitor.v1.alerts.list({ logLevel: "error", limit: 10 });
    result.alerts = alerts.map(alert => ({
      sid: maskSid(alert.sid),
      errorCode: alert.errorCode,
      alertText: alert.alertText,
      resourceSid: maskSid(alert.resourceSid),
      dateGenerated: alert.dateGenerated
    }));
  } catch (err) {
    result.errors.alerts = maskTwilioError(err);
  }

  return result;
}

function formatPhone(phone) {
  phone = String(phone || "").replace(/\s+/g, "");

  if(!phone) return "";

  if (phone.startsWith("+")) return phone;

  if (phone.startsWith("0")) {
    phone = phone.substring(1);
  }

  return "+91" + phone;
}

function normalizeWhatsAppAddress(value, fallback) {
  const raw = String(value || fallback || "").trim().replace(/\s+/g, "");
  if(!raw) return "";

  return raw.startsWith("whatsapp:") ? raw : `whatsapp:${raw}`;
}

function firstEnv(...names) {
  for(const name of names){
    const value = String(process.env[name] || "").trim();
    if(value) return value;
  }

  return "";
}

function getTwilioStatusCallbackUrl() {
  if(!BACKEND_PUBLIC_URL) return "";

  return `${BACKEND_PUBLIC_URL}/twilio-status`;
}

function maskSid(sid) {
  const value = String(sid || "");
  if(!value) return "";

  return `${value.slice(0, 4)}...${value.slice(-6)}`;
}

function maskPhone(value) {
  const raw = String(value || "");
  if(!raw) return "";

  return raw.replace(/(\+?\d{2,4})\d+(\d{4})/g, "$1***$2");
}

function maskTwilioError(err) {
  return {
    code: err.code,
    status: err.status,
    message: err.message
  };
}

function formatTwilioMessage(message) {
  return {
    sid: maskSid(message.sid),
    status: message.status,
    errorCode: message.errorCode,
    errorMessage: message.errorMessage,
    direction: message.direction,
    from: maskPhone(message.from),
    to: maskPhone(message.to),
    dateCreated: message.dateCreated,
    dateSent: message.dateSent
  };
}

function formatTwilioSender(sender) {
  return {
    sid: maskSid(sender.sid),
    senderId: maskPhone(sender.senderId),
    status: sender.status,
    profileName: sender.profileName,
    wabaId: sender.wabaId,
    offlineReasons: sender.offlineReasons,
    dateCreated: sender.dateCreated,
    dateUpdated: sender.dateUpdated
  };
}

function formatTwilioContentApproval(content) {
  return {
    sid: maskSid(content.sid),
    friendlyName: content.friendlyName,
    language: content.language,
    approvalRequests: content.approvalRequests,
    dateCreated: content.dateCreated,
    dateUpdated: content.dateUpdated
  };
}

function getTwilioTemplateDiagnostics() {
  return Object.fromEntries(
    Object.entries(TWILIO_TEMPLATES).map(([name, sid]) => [
      name,
      {
        configured: Boolean(sid),
        sid: maskSid(sid)
      }
    ])
  );
}

function getTestTemplateOptions(template) {
  const templates = {
    reminder: {
      contentSid: TWILIO_TEMPLATES.reminder,
      variables: {
        "1": "Soham",
        "2": "3598",
        "3": APP_PUBLIC_URL
      }
    },
    dueAdded: {
      contentSid: TWILIO_TEMPLATES.dueAdded,
      variables: {
        "1": "Soham",
        "2": "June 2026",
        "3": "1800",
        "4": APP_PUBLIC_URL
      }
    },
    monthlyDue: {
      contentSid: TWILIO_TEMPLATES.monthlyDue,
      variables: {
        "1": "Soham",
        "2": "June 2026",
        "3": "1800",
        "4": APP_PUBLIC_URL
      }
    },
    paymentReceived: {
      contentSid: TWILIO_TEMPLATES.paymentReceived,
      variables: {
        "1": "Soham",
        "2": "1800",
        "3": "0",
        "4": APP_PUBLIC_URL
      }
    }
  };

  const options = templates[template];
  if(!options || !options.contentSid) return null;

  return options;
}

function verifyTestSecret(req, res) {
  const providedSecret = req.headers["x-test-secret"] || req.query.secret;

  if(!process.env.TWILIO_TEST_SECRET || providedSecret !== process.env.TWILIO_TEST_SECRET){
    res.status(404).send("Not found");
    return false;
  }

  return true;
}

function getBillUrl(fileName) {
  if(!BACKEND_PUBLIC_URL || !fileName) return "";
  return `${BACKEND_PUBLIC_URL}/bills/${encodeURIComponent(fileName)}`;
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
    if (payment.description) {
      doc.text(`Description: ${payment.description}`);
    }
    if (payment.paymentMethod) {
      doc.text(`Payment Method: ${payment.paymentMethod}`);
    }
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
    mailTransport: getMailTransport(),
    mailFrom: getMailTransport() === "resend" ? getMailFrom() : null,
    mailRelay: getMailTransport() === "relay" ? getMailRelayUrl() : null,
    razorpayConfigured: Boolean(process.env.RAZORPAY_KEY_ID && process.env.RAZORPAY_KEY_SECRET),
    razorpayMode: getRazorpayMode()
  });
});

app.get("/health-email", async (req, res) => {
  const via = getMailTransport();

  try {
    if(via === "resend"){
      // Cheapest way to prove the key works without sending anything.
      const response = await fetch("https://api.resend.com/domains", {
        headers: { Authorization: `Bearer ${getResendApiKey()}` },
        signal: AbortSignal.timeout(15000)
      });

      if(!response.ok){
        const details = await response.text();

        // A sending-only key is not allowed to list domains. That is the key
        // working exactly as scoped, not a broken key — and sending-only is
        // the scope this server should be using.
        if(!details.includes("restricted_api_key")){
          throw new Error(`Resend key rejected (${response.status}): ${details}`);
        }

        return res.json({
          ok: true,
          via,
          keyScope: "sending-only",
          from: getMailFrom()
        });
      }
    }else if(via === "relay"){
      await verifyMailRelay();
    }else{
      const transporter = await createMailTransporter();
      await transporter.verify();
    }

    res.json({
      ok: true,
      via,
      from: via === "resend" ? getMailFrom() : maskEmail(process.env.EMAIL_USER)
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      via,
      from: via === "resend" ? getMailFrom() : maskEmail(process.env.EMAIL_USER),
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
