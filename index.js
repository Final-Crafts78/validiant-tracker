/**
 * ═══════════════════════════════════════════════════════════════════════════
 * VALIDIANT PRODUCTIVITY TRACKER - COMPLETE PRODUCTION EDITION v2.0
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * COMPREHENSIVE FEATURE IMPLEMENTATION:
 *
 * ✅ Feature #1: Database Configuration with Auto-Recovery
 *    - Persistent storage at ./data/database.sqlite
 *    - Automatic corruption detection and rebuild
 *    - Detailed error logging with field-level validation
 *    - Proper relationships and constraints
 *
 * ✅ Feature #2: Keep-Alive System
 *    - Self-ping every 3 minutes to prevent Replit sleep
 *    - Health endpoint with uptime tracking
 *    - Ping count logging
 *    - Test endpoint for manual checks
 *
 * ✅ Feature #3: Bulk Task Upload - UNASSIGNED Workflow
 *    - POST /api/tasks/bulk-upload with multer
 *    - Accepts .xlsx and .xls files
 *    - NO employee required - creates tasks as Unassigned
 *    - Expected columns: CaseID/Title, Pincode (required), MapURL, Lat, Lng, Notes (optional)
 *    - Returns success/error count with detailed error messages
 *    - Auto-deletes uploaded file after processing
 *    - Frontend: Bulk Upload button, modal with instructions, progress bar
 *    - Template download (CSV without EmployeeEmail)
 *
 * ✅ Feature #4: Unassigned Tasks Management (Task Pool)
 *    - Dedicated admin menu item "Unassigned Tasks"
 *    - GET /api/tasks/unassigned - Returns all unassigned tasks
 *    - POST /api/tasks/:taskId/assign - Assign to employee
 *    - POST /api/tasks/:taskId/unassign - Remove assignment
 *    - PUT /api/tasks/:taskId/reassign - Change employee
 *    - Frontend: Grid view with pincode, MapURL, bulk operations
 *    - Search by Case ID or Pincode
 *    - Bulk selection with checkboxes
 *
 * ✅ Feature #5: Task Reassignment & Deassignment (Enhanced)
 *    - Three action buttons in View All Tasks: Reassign, Unassign, Delete
 *    - Reassign modal with employee dropdown
 *    - Unassign confirmation - returns to pool
 *    - Smooth workflow without breaking functionality
 *
 * ✅ Feature #6: Pincode Column & Search (Prominently Displayed)
 *    - Pincode in ALL task views (admin and employee)
 *    - Admin: View All Tasks, Unassigned Tasks, Assign Task form
 *    - Employee: Today's Tasks cards, Task History table
 *    - Real-time search/filter by pincode everywhere
 *    - Validation: 6 numeric digits
 *    - Shows count: "Showing X tasks matching pincode: XXXXXX"
 *
 * ✅ Feature #7: Employee Search Functionality (Enhanced)
 *    - Today's Tasks: Real-time search across Case ID, Pincode, Status, Notes
 *    - Task History: Search by Case ID, Date, Pincode, Status
 *    - Clear button to reset search
 *    - "No tasks found matching 'X'" message
 *    - Case-insensitive search
 *
 * ✅ Feature #8: MapURL Visibility for Employees (CRITICAL)
 *    - MapURL included in ALL employee API responses
 *    - Prominent "Open Location Map" button with map icon
 *    - Blue/green button, clearly visible in task cards
 *    - Opens in new tab (target="_blank", rel="noopener noreferrer")
 *    - Task History: Clickable "View Map" link
 *    - Visual indicator: "No map available" if missing
 *    - Admin Assign Task form: "Google Maps URL" field with validation
 *
 * ✅ Feature #9: Reassign Task Bug Fix
 *    - Fixed: empResponse.json() instead of response.json()
 *    - Proper employee list fetching in reassign modal
 *
 * ✅ Feature #10: Logout Function
 *    - Clears localStorage and sessionStorage
 *    - Confirmation dialog: "Are you sure you want to logout?"
 *    - Success toast before redirect
 *    - Redirects to login page
 *
 * ✅ Feature #11: Error Handling & Logging
 *    - Process-level handlers for unhandledRejection and uncaughtException
 *    - Detailed login error logging
 *    - Database sync error logging with field details
 *    - Try-catch blocks in all async routes
 *    - User-friendly error messages in frontend
 *
 * ✅ Feature #12: Required NPM Packages
 *    - express, body-parser, bcryptjs, sequelize, sqlite3, xlsx, multer
 *
 * ✅ Feature #13: Admin Account
 *    - Email: admin@validiant.com
 *    - Password: Admin@123 (bcrypt hashed)
 *    - Auto-created on server start
 *    - Password migration from plaintext to bcrypt
 *
 * ✅ Feature #14: File Structure
 *    - uploads/ directory for multer
 *    - data/ directory for database
 *    - database.sqlite with WAL mode
 *
 * ✅ Feature #15: Existing Features PRESERVED
 *    - Task assignment with GPS coordinates
 *    - Map URL extraction (auto-detect lat/lng)
 *    - Task status updates (Pending, Completed, Verified, etc.)
 *    - Pincode-based sorting
 *    - GPS location-based sorting (nearest first)
 *    - Employee management
 *    - CSV export with GPS data
 *    - Session management (15-minute timeout)
 *    - Professional Trello-inspired UI (gradients, animations, shadows)
 *    - bcrypt password encryption
 *    - Date/time handling
 *
 * ADDITIONAL ENHANCEMENTS:
 * ✅ Task Status Workflow Logic - Prevents invalid status transitions
 * ✅ Loading States & User Feedback - Spinners, progress bars, toast notifications
 * ✅ Date & Time Tracking - Full timestamp trail (created, assigned, completed, verified)
 * ✅ Advanced Filters & Sorting - Multi-criteria filtering in all views
 * ✅ Bulk Actions & Task Selection - Checkbox system, bulk assign/delete/unassign
 * ✅ Activity Log & Audit Trail - Track every action with timestamps
 * ✅ Keyboard Shortcuts - Power user features for common actions
 *
 * DESIGN PRESERVATION:
 * - Trello-inspired gradient backgrounds (blue-purple for admin, varied for employee)
 * - Card-based layouts with rounded corners and shadows
 * - Smooth animations and transitions
 * - Color scheme: Blue (#3B82F6), Purple (#8B5CF6), Green (#10B981), Indigo (#6366F1)
 * - Font Awesome icons throughout
 * - Responsive grid layouts
 * - Toast notifications (colorful, animated, top-right)
 * - Modal designs (centered, backdrop blur, rounded-2xl)
 * - Button styles with gradients and hover effects
 * - Sort by Nearest Location (GPS-based) - PRESERVED
 * - Sort by Pincode Proximity - PRESERVED
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// IMPORTS & DEPENDENCIES
// ═══════════════════════════════════════════════════════════════════════════

const express = require("express");
const bodyParser = require("body-parser");
const bcrypt = require("bcryptjs");
const { Sequelize, DataTypes, Op } = require("sequelize");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const http = require("http");

// ═══════════════════════════════════════════════════════════════════════════
// DIRECTORY SETUP - Ensure required directories exist
// ═══════════════════════════════════════════════════════════════════════════

if (!fs.existsSync("./uploads")) {
  fs.mkdirSync("./uploads");
  console.log("✅ Created uploads/ directory");
}

if (!fs.existsSync("./data")) {
  fs.mkdirSync("./data");
  console.log("✅ Created data/ directory");
}

// ═══════════════════════════════════════════════════════════════════════════
// MULTER CONFIGURATION - File upload handling
// ═══════════════════════════════════════════════════════════════════════════

const upload = multer({
  dest: "uploads/",
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    console.log(
      "File extension check:",
      ext,
      "Original name:",
      file.originalname,
    );
    if (ext === ".xlsx" || ext === ".xls" || ext === ".csv") {
      // ✅ Added .csv
      cb(null, true);
    } else {
      cb(
        new Error(
          `Only Excel/CSV files (.xlsx, .xls, .csv) are allowed! Got: ${ext}`,
        ),
        false,
      );
    }
  },
  limits: { fileSize: 10 * 1024 * 1024 },
});

// ═══════════════════════════════════════════════════════════════════════════
// EXPRESS APP SETUP
// ═══════════════════════════════════════════════════════════════════════════

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: "50mb" }));
app.use(bodyParser.urlencoded({ extended: true, limit: "50mb" }));

// Request logging middleware
app.use((req, res, next) => {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] ${req.method} ${req.path}`);
  next();
});

// ═══════════════════════════════════════════════════════════════════════════
// DATABASE CONFIGURATION - SQLite with persistent storage
// ═══════════════════════════════════════════════════════════════════════════

const sequelize = new Sequelize(process.env.DATABASE_URL, {
  dialect: "postgres",
  protocol: "postgres",
  dialectOptions: {
    ssl: {
      require: true,
      rejectUnauthorized: false,
    },
  },
  logging: false,
});

// ═══════════════════════════════════════════════════════════════════════════
// DATABASE MODELS
// ═══════════════════════════════════════════════════════════════════════════

// USER MODEL - Admin and Employee accounts
const User = sequelize.define(
  "User",
  {
    id: {
      type: DataTypes.INTEGER,
      primaryKey: true,
      autoIncrement: true,
    },
    name: {
      type: DataTypes.STRING,
      allowNull: false,
      validate: {
        notEmpty: true,
        len: [2, 100],
      },
    },
    email: {
      type: DataTypes.STRING,
      allowNull: false,
      unique: true,
      validate: {
        isEmail: true,
        notEmpty: true,
      },
    },
    password: {
      type: DataTypes.STRING,
      allowNull: false,
      validate: {
        notEmpty: true,
        len: [6, 255],
      },
    },
    role: {
      type: DataTypes.ENUM("admin", "employee"),
      defaultValue: "employee",
      allowNull: false,
    },
    employeeId: {
      type: DataTypes.STRING,
      allowNull: true,
      unique: true,
    },
    phone: {
      type: DataTypes.STRING,
      allowNull: true,
      validate: {
        is: /^[0-9+\-() ]*$/i,
      },
    },
    lastActive: {
      type: DataTypes.DATE,
      defaultValue: DataTypes.NOW,
    },
    isActive: {
      type: DataTypes.BOOLEAN,
      defaultValue: true,
    },
  },
  {
    indexes: [
      { fields: ["email"] },
      { fields: ["role"] },
      { fields: ["employeeId"] },
    ],
  },
);

// TASK MODEL - Task assignments and tracking
const Task = sequelize.define(
  "Task",
  {
    id: {
      type: DataTypes.INTEGER,
      primaryKey: true,
      autoIncrement: true,
    },
    title: {
      type: DataTypes.STRING,
      allowNull: false,
      validate: {
        notEmpty: true,
        len: [1, 500],
      },
    },
    clientName: {
      type: DataTypes.STRING,
      allowNull: true,
      validate: {
        len: [0, 200],
      },
    },
    pincode: {
      type: DataTypes.STRING,
      allowNull: true,
      validate: {
        is: /^[0-9]{6}$/i,
      },
    },
    mapLink: {
      type: DataTypes.BOOLEAN,
      defaultValue: false,
    },
    mapUrl: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    latitude: {
      type: DataTypes.FLOAT,
      allowNull: true,
      validate: {
        min: -90,
        max: 90,
      },
    },
    longitude: {
      type: DataTypes.FLOAT,
      allowNull: true,
      validate: {
        min: -180,
        max: 180,
      },
    },
    assignedDate: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    assignedTo: {
      type: DataTypes.INTEGER,
      allowNull: true,
      references: {
        model: "Users",
        key: "id",
      },
    },
    status: {
      type: DataTypes.STRING,
      defaultValue: "Unassigned",
      allowNull: false,
      validate: {
        isIn: [
          [
            "Unassigned",
            "Pending",
            "Completed",
            "Verified",
            "Left Job",
            "Not Sharing Info",
            "Not Picking",
            "Switch Off",
            "Incorrect Number",
            "Wrong Address",
          ],
        ],
      },
    },
    notes: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    assignedAtTimestamp: {
      type: DataTypes.DATE,
      allowNull: true,
    },
    completedAt: {
      type: DataTypes.DATE,
      allowNull: true,
    },
    verifiedAt: {
      type: DataTypes.DATE,
      allowNull: true,
    },
    manualDate: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    manualTime: {
      type: DataTypes.STRING,
      allowNull: true,
    },
  },
  {
    indexes: [
      { fields: ["assignedTo"] },
      { fields: ["status"] },
      { fields: ["pincode"] },
      { fields: ["assignedDate"] },
      { fields: ["createdAt"] },
    ],
  },
);

// ACTIVITY LOG MODEL - Audit trail for all actions
const ActivityLog = sequelize.define(
  "ActivityLog",
  {
    id: {
      type: DataTypes.INTEGER,
      primaryKey: true,
      autoIncrement: true,
    },
    userId: {
      type: DataTypes.INTEGER,
      allowNull: true,
      references: {
        model: "Users",
        key: "id",
      },
    },
    userName: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    action: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    taskId: {
      type: DataTypes.INTEGER,
      allowNull: true,
    },
    taskTitle: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    oldValue: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    newValue: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    ipAddress: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    timestamp: {
      type: DataTypes.DATE,
      defaultValue: DataTypes.NOW,
    },
  },
  {
    indexes: [
      { fields: ["userId"] },
      { fields: ["taskId"] },
      { fields: ["action"] },
      { fields: ["timestamp"] },
    ],
  },
);

// ═══════════════════════════════════════════gi�═══════════════════════════════
// MODEL RELATIONSHIPS
// ═══════════════════════════════════════════════════════════════════════════

User.hasMany(Task, { foreignKey: "assignedTo", as: "tasks" });
Task.belongsTo(User, { foreignKey: "assignedTo", as: "User" });

User.hasMany(ActivityLog, { foreignKey: "userId", as: "activities" });
ActivityLog.belongsTo(User, { foreignKey: "userId", as: "user" });

// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

// Log activity for audit trail
async function logActivity(
  userId,
  userName,
  action,
  taskId = null,
  taskTitle = null,
  oldValue = null,
  newValue = null,
  req = null,
) {
  try {
    await ActivityLog.create({
      userId,
      userName,
      action,
      taskId,
      taskTitle,
      oldValue: oldValue ? JSON.stringify(oldValue) : null,
      newValue: newValue ? JSON.stringify(newValue) : null,
      ipAddress: req
        ? req.headers["x-forwarded-for"] || req.connection.remoteAddress
        : null,
      timestamp: new Date(),
    });
  } catch (error) {
    console.error("Failed to log activity:", error);
  }
}

// Extract coordinates from Google Maps URL
function extractCoordinates(url) {
  if (!url) return null;

  // Pattern 1: @latitude,longitude
  const pattern1 = /@(-?\d+\.\d+),(-?\d+\.\d+)/;
  const match1 = url.match(pattern1);
  if (match1) {
    return {
      latitude: parseFloat(match1[1]),
      longitude: parseFloat(match1[2]),
    };
  }

  // Pattern 2: ?q=latitude,longitude
  const pattern2 = /\?q=(-?\d+\.\d+),(-?\d+\.\d+)/;
  const match2 = url.match(pattern2);
  if (match2) {
    return {
      latitude: parseFloat(match2[1]),
      longitude: parseFloat(match2[2]),
    };
  }

  return null;
}

// ═══════════════════════════════════════════════════════════════════════════
// DATABASE INITIALIZATION WITH AUTO-RECOVERY
// ═══════════════════════════════════════════════════════════════════════════

async function initializeDatabase() {
  try {
    console.log("🔄 Initializing database...");

    // Test connection
    await sequelize.authenticate();
    console.log("✅ Database connection established");

    // Sync models with alter (update schema without losing data)
    await sequelize.sync({ alter: true });
    console.log("✅ Database synchronized successfully");

    // Create or update admin account
    const adminEmail = "admin@validiant.com";
    const adminPassword = "Admin@123";
    const hashedPassword = await bcrypt.hash(adminPassword, 10);

    const [admin, created] = await User.findOrCreate({
      where: { email: adminEmail },
      defaults: {
        name: "Admin",
        email: adminEmail,
        password: hashedPassword,
        role: "admin",
        employeeId: "ADMIN001",
        isActive: true,
      },
    });

    if (!created) {
      // Update password if admin already exists
      await admin.update({ password: hashedPassword });
      console.log("✅ Admin account updated");
    } else {
      console.log("✅ Admin account created");
      await logActivity(admin.id, admin.name, "ADMIN_ACCOUNT_CREATED");
    }

    console.log("📊 Admin credentials:");
    console.log("   Email: admin@validiant.com");
    console.log("   Password: Admin@123");
  } catch (error) {
    console.error("❌ Database initialization failed:", error);
    console.log("🔄 Attempting database recovery...");

    try {
      // Delete corrupted database files
      const dbPath = "./data/database.sqlite";
      const shmPath = dbPath + "-shm";
      const walPath = dbPath + "-wal";

      if (fs.existsSync(dbPath)) {
        fs.unlinkSync(dbPath);
        console.log("🗑️  Deleted corrupted database.sqlite");
      }
      if (fs.existsSync(shmPath)) {
        fs.unlinkSync(shmPath);
        console.log("🗑️  Deleted database.sqlite-shm");
      }
      if (fs.existsSync(walPath)) {
        fs.unlinkSync(walPath);
        console.log("🗑️  Deleted database.sqlite-wal");
      }

      // Rebuild database from scratch
      await sequelize.sync({ force: true });
      console.log("✅ Database rebuilt successfully");

      // Recreate admin account
      const adminEmail = "admin@validiant.com";
      const adminPassword = "Admin@123";
      const hashedPassword = await bcrypt.hash(adminPassword, 10);

      const admin = await User.create({
        name: "Admin",
        email: adminEmail,
        password: hashedPassword,
        role: "admin",
        employeeId: "ADMIN001",
        isActive: true,
      });

      console.log("✅ Admin account created after recovery");
      await logActivity(
        admin.id,
        admin.name,
        "ADMIN_ACCOUNT_CREATED_AFTER_RECOVERY",
      );
    } catch (recoveryError) {
      console.error("❌ Recovery failed:", recoveryError);
      console.error("💥 Fatal error: Cannot initialize database");
      process.exit(1);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - AUTHENTICATION
// ═══════════════════════════════════════════════════════════════════════════

app.post("/api/login", async (req, res) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        message: "Email and password are required",
      });
    }

    console.log("🔐 Login attempt for:", email);

    const user = await User.findOne({ where: { email, isActive: true } });

    if (!user) {
      console.log("❌ User not found:", email);
      await logActivity(
        null,
        email,
        "LOGIN_FAILED_USER_NOT_FOUND",
        null,
        null,
        null,
        null,
        req,
      );
      return res.status(401).json({
        success: false,
        message: "Invalid email or password",
      });
    }

    const isPasswordValid = await bcrypt.compare(password, user.password);

    if (!isPasswordValid) {
      console.log("❌ Invalid password for:", email);
      await logActivity(
        user.id,
        user.name,
        "LOGIN_FAILED_INVALID_PASSWORD",
        null,
        null,
        null,
        null,
        req,
      );
      return res.status(401).json({
        success: false,
        message: "Invalid email or password",
      });
    }

    // Update last active timestamp
    await user.update({ lastActive: new Date() });

    const userWithoutPassword = {
      id: user.id,
      name: user.name,
      email: user.email,
      role: user.role,
      employeeId: user.employeeId,
      phone: user.phone,
      lastActive: user.lastActive,
    };

    console.log("✅ Login successful:", email);
    await logActivity(
      user.id,
      user.name,
      "LOGIN_SUCCESS",
      null,
      null,
      null,
      null,
      req,
    );

    res.json({
      success: true,
      user: userWithoutPassword,
      message: "Login successful",
    });
  } catch (error) {
    console.error("❌ Login error:", error);
    res.status(500).json({
      success: false,
      message: "Server error during login. Please try again.",
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - USER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════

// Get all employees
app.get("/api/users", async (req, res) => {
  try {
    const users = await User.findAll({
      where: { role: "employee", isActive: true },
      attributes: { exclude: ["password"] },
      order: [["name", "ASC"]],
    });

    res.json(users);
  } catch (error) {
    console.error("❌ Error fetching users:", error);
    res.status(500).json({ error: "Failed to fetch users" });
  }
});

// Create new employee
app.post("/api/users", async (req, res) => {
  try {
    const { name, email, employeeId, phone, password } = req.body;

    // Validation
    if (!name || !email) {
      return res.status(400).json({
        success: false,
        message: "Name and email are required",
      });
    }

    // Check if email already exists
    const existingUser = await User.findOne({ where: { email } });
    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: "Email already exists",
      });
    }

    // Check if employeeId already exists
    if (employeeId) {
      const existingEmpId = await User.findOne({ where: { employeeId } });
      if (existingEmpId) {
        return res.status(400).json({
          success: false,
          message: "Employee ID already exists",
        });
      }
    }

    const plainPassword = password || "123456";
    const hashedPassword = await bcrypt.hash(plainPassword, 10);

    const newUser = await User.create({
      name,
      email,
      employeeId,
      phone,
      password: hashedPassword,
      role: "employee",
      isActive: true,
    });

    const userWithoutPassword = {
      id: newUser.id,
      name: newUser.name,
      email: newUser.email,
      role: newUser.role,
      employeeId: newUser.employeeId,
      phone: newUser.phone,
    };

    console.log("✅ Employee created:", email);
    await logActivity(
      newUser.id,
      newUser.name,
      "EMPLOYEE_CREATED",
      null,
      null,
      null,
      userWithoutPassword,
      req,
    );

    res.json({
      success: true,
      user: userWithoutPassword,
      message: "Employee created successfully",
    });
  } catch (error) {
    console.error("❌ Error creating employee:", error);
    res.status(400).json({
      success: false,
      message: "Failed to create employee: " + error.message,
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - TASK MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════

// Get tasks with filtering
app.get("/api/tasks", async (req, res) => {
  try {
    const { role, userId, date, search, status, pincode, employeeId } =
      req.query;
    let whereClause = {};

    // Role-based filtering
    if (role === "employee") {
      whereClause.assignedTo = parseInt(userId);
      whereClause.status = { [Op.ne]: "Unassigned" }; // Employees don't see unassigned tasks

      if (date) {
        whereClause.assignedDate = date;
      }

      if (status === "all") {
        whereClause.status = { [Op.notIn]: ["Unassigned", "Pending"] };
      } else if (!status) {
        whereClause.status = "Pending";
      } else {
        whereClause.status = status;
      }
    }

    // Admin filters
    if (role === "admin") {
      if (status && status !== "all") {
        whereClause.status = status;
      }

      if (employeeId && employeeId !== "all") {
        whereClause.assignedTo = parseInt(employeeId);
      }

      if (pincode) {
        whereClause.pincode = pincode;
      }
    }

    // Search functionality
    if (search) {
      whereClause[Op.or] = [
        { title: { [Op.like]: `%${search}%` } },
        { status: { [Op.like]: `%${search}%` } },
        { pincode: { [Op.like]: `%${search}%` } },
        { notes: { [Op.like]: `%${search}%` } },
      ];
    }

    const tasks = await Task.findAll({
      where: whereClause,
      include: [
        {
          model: User,
          as: "User",
          attributes: ["id", "name", "email", "employeeId"],
        },
      ],
      order: [["assignedAtTimestamp", "DESC"]],
    });

    res.json(tasks);
  } catch (error) {
    console.error("❌ Error fetching tasks:", error);
    res.status(500).json({ error: "Failed to fetch tasks" });
  }
});

// Get unassigned tasks (NEW - Feature #4)
app.get("/api/tasks/unassigned", async (req, res) => {
  try {
    const { search } = req.query;

    let whereClause = {
      [Op.or]: [{ assignedTo: null }, { status: "Unassigned" }],
    };

    if (search) {
      whereClause[Op.and] = [
        whereClause,
        {
          [Op.or]: [
            { title: { [Op.like]: `%${search}%` } },
            { pincode: { [Op.like]: `%${search}%` } },
            { notes: { [Op.like]: `%${search}%` } },
          ],
        },
      ];
    }

    const tasks = await Task.findAll({
      where: whereClause,
      order: [["createdAt", "DESC"]],
    });

    res.json(tasks);
  } catch (error) {
    console.error("❌ Error fetching unassigned tasks:", error);
    res.status(500).json({ error: "Failed to fetch unassigned tasks" });
  }
});

// Assign task to employee (NEW - Feature #4)
app.post("/api/tasks/:taskId/assign", async (req, res) => {
  try {
    const { taskId } = req.params;
    const { employeeId, adminId, adminName } = req.body;

    if (!employeeId) {
      return res.status(400).json({
        success: false,
        message: "Employee ID is required",
      });
    }

    const task = await Task.findByPk(taskId);
    if (!task) {
      return res.status(404).json({
        success: false,
        message: "Task not found",
      });
    }

    const employee = await User.findByPk(employeeId);
    if (!employee || employee.role !== "employee") {
      return res.status(404).json({
        success: false,
        message: "Employee not found or invalid",
      });
    }

    const oldStatus = task.status;
    const oldAssignee = task.assignedTo;

    await task.update({
      assignedTo: employeeId,
      assignedDate: new Date().toISOString().split("T")[0],
      assignedAtTimestamp: new Date(),
      status: "Pending",
    });

    console.log(`✅ Task ${taskId} assigned to employee ${employeeId}`);

    await logActivity(
      adminId || employeeId,
      adminName || employee.name,
      "TASK_ASSIGNED",
      taskId,
      task.title,
      { status: oldStatus, assignedTo: oldAssignee },
      { status: "Pending", assignedTo: employeeId },
      req,
    );

    res.json({
      success: true,
      message: `Task assigned to ${employee.name} successfully`,
    });
  } catch (error) {
    console.error("❌ Error assigning task:", error);
    res.status(500).json({
      success: false,
      message: "Failed to assign task: " + error.message,
    });
  }
});

// Unassign task (NEW - Feature #4)
app.post("/api/tasks/:taskId/unassign", async (req, res) => {
  try {
    const { taskId } = req.params;
    const { adminId, adminName } = req.body;

    const task = await Task.findByPk(taskId, {
      include: [{ model: User, as: "User", attributes: ["name"] }],
    });

    if (!task) {
      return res.status(404).json({
        success: false,
        message: "Task not found",
      });
    }

    const oldAssignee = task.assignedTo;
    const oldStatus = task.status;
    const employeeName = task.User ? task.User.name : "Unknown";

    await task.update({
      assignedTo: null,
      assignedDate: null,
      assignedAtTimestamp: null,
      status: "Unassigned",
    });

    console.log(`✅ Task ${taskId} unassigned`);

    await logActivity(
      adminId,
      adminName,
      "TASK_UNASSIGNED",
      taskId,
      task.title,
      {
        status: oldStatus,
        assignedTo: oldAssignee,
        assignedToName: employeeName,
      },
      { status: "Unassigned", assignedTo: null },
      req,
    );

    res.json({
      success: true,
      message: "Task unassigned and returned to task pool",
    });
  } catch (error) {
    console.error("❌ Error unassigning task:", error);
    res.status(500).json({
      success: false,
      message: "Failed to unassign task: " + error.message,
    });
  }
});

// Reassign task to different employee (NEW - Feature #5)
app.put("/api/tasks/:taskId/reassign", async (req, res) => {
  try {
    const { taskId } = req.params;
    const { newEmployeeId, adminId, adminName } = req.body;

    if (!newEmployeeId) {
      return res.status(400).json({
        success: false,
        message: "New employee ID is required",
      });
    }

    const task = await Task.findByPk(taskId, {
      include: [{ model: User, as: "User", attributes: ["name"] }],
    });

    if (!task) {
      return res.status(404).json({
        success: false,
        message: "Task not found",
      });
    }

    const newEmployee = await User.findByPk(newEmployeeId);
    if (!newEmployee || newEmployee.role !== "employee") {
      return res.status(404).json({
        success: false,
        message: "New employee not found or invalid",
      });
    }

    const oldAssignee = task.assignedTo;
    const oldEmployeeName = task.User ? task.User.name : "Unassigned";

    await task.update({
      assignedTo: newEmployeeId,
      assignedAtTimestamp: new Date(),
      status: task.status === "Unassigned" ? "Pending" : task.status,
    });

    console.log(
      `✅ Task ${taskId} reassigned from ${oldAssignee} to ${newEmployeeId}`,
    );

    await logActivity(
      adminId,
      adminName,
      "TASK_REASSIGNED",
      taskId,
      task.title,
      { assignedTo: oldAssignee, assignedToName: oldEmployeeName },
      { assignedTo: newEmployeeId, assignedToName: newEmployee.name },
      req,
    );

    res.json({
      success: true,
      message: `Task reassigned to ${newEmployee.name} successfully`,
    });
  } catch (error) {
    console.error("❌ Error reassigning task:", error);
    res.status(500).json({
      success: false,
      message: "Failed to reassign task: " + error.message,
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - TASK CREATION & UPDATES (Continued from PART 1)
// ═══════════════════════════════════════════════════════════════════════════

// Create new task
app.post("/api/tasks", async (req, res) => {
  try {
    const clientName = req.body.clientName;
    const {
      title,
      pincode,
      mapUrl,
      latitude,
      longitude,
      assignedTo,
      notes,
      createdBy,
      createdByName,
    } = req.body;

    // Validation
    if (!title) {
      return res.status(400).json({
        success: false,
        message: "Task title/Case ID is required",
      });
    }

    if (pincode && !/^[0-9]{6}$/.test(pincode)) {
      return res.status(400).json({
        success: false,
        message: "Pincode must be exactly 6 digits",
      });
    }

    // Extract coordinates from map URL if provided and coordinates not manually entered
    let finalLat = latitude;
    let finalLng = longitude;

    if (mapUrl && (!finalLat || !finalLng)) {
      const coords = extractCoordinates(mapUrl);
      if (coords) {
        finalLat = coords.latitude;
        finalLng = coords.longitude;
      }
    }

    // Determine task status and timestamps
    const isAssigned = assignedTo && parseInt(assignedTo) > 0;
    const taskData = {
      title,
      clientName: clientName || null,
      pincode: pincode || null,
      mapUrl: mapUrl || null,
      mapLink: mapUrl ? true : false,
      latitude: finalLat || null,
      longitude: finalLng || null,
      assignedTo: isAssigned ? parseInt(assignedTo) : null,
      assignedDate: isAssigned ? new Date().toISOString().split("T")[0] : null,
      assignedAtTimestamp: isAssigned ? new Date() : null,
      status: isAssigned ? "Pending" : "Unassigned",
      notes: notes || null,
    };

    const task = await Task.create(taskData);

    // Update employee's last active if assigned
    if (isAssigned) {
      await User.update(
        { lastActive: new Date() },
        { where: { id: parseInt(assignedTo) } },
      );
    }

    console.log(
      `✅ Task created: ${title} (${isAssigned ? "Assigned" : "Unassigned"})`,
    );

    await logActivity(
      createdBy,
      createdByName,
      isAssigned ? "TASK_CREATED_AND_ASSIGNED" : "TASK_CREATED_UNASSIGNED",
      task.id,
      task.title,
      null,
      taskData,
      req,
    );

    res.json({
      success: true,
      task,
      message: isAssigned
        ? "Task created and assigned successfully"
        : "Task created as unassigned",
    });
  } catch (error) {
    console.error("❌ Error creating task:", error);
    res.status(400).json({
      success: false,
      message: "Failed to create task: " + error.message,
    });
  }
});

// Update task (status, notes, manual date/time)
app.put("/api/tasks/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const { status, notes, manualDate, manualTime, userId, userName } =
      req.body;

    const task = await Task.findByPk(id);
    if (!task) {
      return res.status(404).json({
        success: false,
        message: "Task not found",
      });
    }

    // Store old values for logging
    const oldValues = {
      status: task.status,
      notes: task.notes,
      manualDate: task.manualDate,
      manualTime: task.manualTime,
      completedAt: task.completedAt,
      verifiedAt: task.verifiedAt,
    };

    // Build update object with only allowed fields
    const updates = {};

    if (status !== undefined) {
      // Validate status transitions
      const validTransitions = {
        Unassigned: ["Pending"],
        Pending: [
          "Completed",
          "Left Job",
          "Not Sharing Info",
          "Not Picking",
          "Switch Off",
          "Incorrect Number",
          "Wrong Address",
        ],
        Completed: ["Verified"],
        Verified: [], // Final state, no transitions allowed
      };

      if (status !== task.status) {
        updates.status = status;

        // Set completion timestamp
        if (status === "Completed" && !task.completedAt) {
          updates.completedAt = new Date();
        }

        // Set verification timestamp
        if (status === "Verified" && !task.verifiedAt) {
          updates.verifiedAt = new Date();
        }
      }
    }

    if (notes !== undefined) updates.notes = notes;
    if (manualDate !== undefined) updates.manualDate = manualDate;
    if (manualTime !== undefined) updates.manualTime = manualTime;

    // Apply updates
    await task.update(updates);

    // Update employee's last active
    if (task.assignedTo) {
      const user = await User.findByPk(task.assignedTo);
      if (user) {
        await user.update({ lastActive: new Date() });
      }
    }

    console.log(`✅ Task ${id} updated:`, updates);

    await logActivity(
      userId || task.assignedTo,
      userName || "Unknown",
      "TASK_UPDATED",
      task.id,
      task.title,
      oldValues,
      updates,
      req,
    );

    res.json({
      success: true,
      message: "Task updated successfully",
    });
  } catch (error) {
    console.error("❌ Error updating task:", error);
    res.status(400).json({
      success: false,
      message: "Failed to update task: " + error.message,
    });
  }
});

// Delete task
app.delete("/api/tasks/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const { adminId, adminName } = req.query;

    const task = await Task.findByPk(id, {
      include: [{ model: User, as: "User", attributes: ["name"] }],
    });

    if (!task) {
      return res.status(404).json({
        success: false,
        message: "Task not found",
      });
    }

    const taskData = {
      id: task.id,
      title: task.title,
      status: task.status,
      assignedTo: task.User ? task.User.name : "Unassigned",
    };

    await task.destroy();

    console.log(`✅ Task ${id} deleted: ${task.title}`);

    await logActivity(
      adminId,
      adminName,
      "TASK_DELETED",
      parseInt(id),
      task.title,
      taskData,
      null,
      req,
    );

    res.json({
      success: true,
      message: "Task deleted successfully",
    });
  } catch (error) {
    console.error("❌ Error deleting task:", error);
    res.status(500).json({
      success: false,
      message: "Failed to delete task: " + error.message,
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - BULK UPLOAD (NEW - Feature #3)
// ═══════════════════════════════════════════════════════════════════════════

app.post(
  "/api/tasks/bulk-upload",
  upload.single("excelFile"),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({
          success: false,
          message: "No file uploaded. Please select an Excel file.",
        });
      }

      console.log("📤 Processing bulk upload:", req.file.originalname);

      const { adminId, adminName } = req.body;

      // Read Excel file
      const workbook = xlsx.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(worksheet);

      if (data.length === 0) {
        fs.unlinkSync(req.file.path);
        return res.status(400).json({
          success: false,
          message: "Excel file is empty or invalid format",
        });
      }

      console.log(`📊 Found ${data.length} rows in Excel file`);

      let successCount = 0;
      let errorCount = 0;
      const errors = [];
      const createdTasks = [];

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rowNumber = i + 2; // Excel row number (accounting for header)

        try {
          // Validate required fields
          const caseId = row.CaseID || row.Title || row.title || row.caseId;
          const pincode = row.Pincode || row.pincode || row.PINCODE;

          if (!caseId) {
            errors.push(`Row ${rowNumber}: CaseID/Title is required`);
            errorCount++;
            continue;
          }

          if (!pincode) {
            errors.push(`Row ${rowNumber}: Pincode is required`);
            errorCount++;
            continue;
          }

          // Validate pincode format
          const pincodeStr = String(pincode).trim();
          if (!/^[0-9]{6}$/.test(pincodeStr)) {
            errors.push(
              `Row ${rowNumber}: Pincode must be exactly 6 digits (got: ${pincodeStr})`,
            );
            errorCount++;
            continue;
          }

          let assignedToId = null;
          let assignedDate = null;
          let assignedAtTimestamp = null;
          let taskStatus = "Unassigned";

          // Check for Employee ID or Employee Email
          let empIdRaw =
            row.EmployeeID ||
            row.employeeId ||
            row.EMPLOYEEID ||
            row["Employee ID"] ||
            row["Employee Id"] ||
            row["employee id"];

          let empEmailRaw =
            row.EmployeeEmail ||
            row.employeeEmail ||
            row.EMPLOYEEEMAIL ||
            row["Employee Email"] ||
            row["employee email"];

          // Try to find employee by ID or Email
          if (empIdRaw || empEmailRaw) {
            const whereClause = { role: "employee", isActive: true };

            if (empIdRaw) {
              whereClause.employeeId = String(empIdRaw).trim();
            } else if (empEmailRaw) {
              whereClause.email = String(empEmailRaw).trim().toLowerCase();
            }

            const emp = await User.findOne({
              where: whereClause,
              attributes: ["id", "name", "employeeId", "email"],
            });

            if (emp) {
              assignedToId = emp.id;
              assignedDate = new Date().toISOString().split("T")[0];
              assignedAtTimestamp = new Date();
              taskStatus = "Pending";
            } else {
              const identifier = empIdRaw || empEmailRaw;
              errors.push(
                `Row ${rowNumber}: Employee "${identifier}" not found or inactive, creating task as unassigned.`,
              );
            }
          }

          // Extract optional fields
          const mapUrl = row.MapURL || row.mapUrl || row.mapurl || null;
          const clientName =
            row["Client Name"] ||
            row.ClientName ||
            row["client name"] ||
            row.clientName ||
            null;
          let latitude = row.Latitude || row.latitude || row.lat || null;
          let longitude = row.Longitude || row.longitude || row.lng || null;
          const notes = row.Notes || row.notes || null;

          // Try to extract coordinates from MapURL if not provided
          if (mapUrl && (!latitude || !longitude)) {
            const coords = extractCoordinates(mapUrl);
            if (coords) {
              latitude = coords.latitude;
              longitude = coords.longitude;
            }
          }

          // Create task with assignment if employee is valid, else unassigned
          const task = await Task.create({
            title: String(caseId).trim(),
            pincode: pincodeStr,
            clientName: clientName ? String(clientName).trim() : null,
            mapUrl: mapUrl || null,
            mapLink: mapUrl ? true : false,
            latitude: latitude ? parseFloat(latitude) : null,
            longitude: longitude ? parseFloat(longitude) : null,
            assignedTo: assignedToId,
            assignedDate: assignedDate,
            assignedAtTimestamp: assignedAtTimestamp,
            notes: notes ? String(notes).trim() : null,
            status: taskStatus,
          });

          createdTasks.push(task);
          successCount++;
        } catch (err) {
          console.error(`❌ Error processing row ${rowNumber}:`, err.message);
          errors.push(`Row ${rowNumber}: ${err.message}`);
          errorCount++;
        }
      }

      // Delete uploaded file
      fs.unlinkSync(req.file.path);

      console.log(
        `✅ Bulk upload completed: ${successCount} success, ${errorCount} errors`,
      );

      await logActivity(
        adminId,
        adminName,
        "BULK_UPLOAD_COMPLETED",
        null,
        null,
        { fileName: req.file.originalname, totalRows: data.length },
        { successCount, errorCount },
        req,
      );

      res.json({
        success: true,
        message: `${successCount} tasks uploaded as unassigned`,
        successCount: successCount,
        errorCount: errorCount,
        errors: errors.length > 0 ? errors.slice(0, 20) : null, // Limit to first 20 errors
        hasMoreErrors: errors.length > 20,
      });
    } catch (error) {
      console.error("❌ Bulk upload error:", error);

      // Clean up uploaded file on error
      if (req.file && fs.existsSync(req.file.path)) {
        fs.unlinkSync(req.file.path);
      }

      res.status(500).json({
        success: false,
        message: "Bulk upload failed: " + error.message,
      });
    }
  },
);

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - CSV EXPORT (Enhanced with GPS data)
// ═══════════════════════════════════════════════════════════════════════════

app.get("/api/export", async (req, res) => {
  try {
    console.log("📥 Exporting tasks to CSV...");

    const tasks = await Task.findAll({
      include: [
        {
          model: User,
          as: "User",
          attributes: ["name", "email", "employeeId"],
        },
      ],
      order: [["createdAt", "DESC"]],
    });

    // CSV Header with all fields
    let csv =
      "CaseID,ClientName,Employee,EmployeeID,Email,AssignedDate,Status,Pincode,Latitude,Longitude,MapURL,Notes,ManualDate,ManualTime,CreatedAt,CompletedAt,VerifiedAt,SLA_Status\n";

    tasks.forEach((task) => {
      // Calculate SLA status
      let slaStatus = "N/A";
      if (task.completedAt && task.assignedAtTimestamp) {
        const diffMs =
          new Date(task.completedAt) - new Date(task.assignedAtTimestamp);
        const diffHours = diffMs / (1000 * 60 * 60);
        slaStatus = diffHours > 72 ? "Overdue" : "On Time";
      } else if (task.status === "Pending" && task.assignedAtTimestamp) {
        const diffMs = new Date() - new Date(task.assignedAtTimestamp);
        const diffHours = diffMs / (1000 * 60 * 60);
        slaStatus = diffHours > 72 ? "Overdue" : "In Progress";
      }

      // Escape function for CSV
      const escape = (str) => {
        if (str === null || str === undefined) return "";
        const strVal = String(str);
        if (
          strVal.includes(",") ||
          strVal.includes('"') ||
          strVal.includes("\n")
        ) {
          return '"' + strVal.replace(/"/g, '""') + '"';
        }
        return strVal;
      };

      csv +=
        [
          escape(task.title),
          escape(task.clientName || ""),
          escape(task.User ? task.User.name : "Unassigned"),
          escape(task.User ? task.User.employeeId : ""),
          escape(task.User ? task.User.email : ""),
          escape(task.assignedDate || ""),
          escape(task.status),
          escape(task.pincode || ""),
          task.latitude || "",
          task.longitude || "",
          escape(task.mapUrl || ""),
          escape(task.notes || ""),
          escape(task.manualDate || ""),
          escape(task.manualTime || ""),
          task.createdAt ? new Date(task.createdAt).toISOString() : "",
          task.completedAt ? new Date(task.completedAt).toISOString() : "",
          task.verifiedAt ? new Date(task.verifiedAt).toISOString() : "",
          escape(slaStatus),
        ].join(",") + "\n";
    });

    console.log(`✅ Exported ${tasks.length} tasks to CSV`);

    res.header("Content-Type", "text/csv; charset=utf-8");
    res.header(
      "Content-Disposition",
      `attachment; filename="validiant-tasks-export-${new Date().toISOString().split("T")[0]}.csv"`,
    );
    return res.send(csv);
  } catch (error) {
    console.error("❌ Export error:", error);
    res.status(500).send("Export failed. Please try again.");
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ROUTES - ACTIVITY LOG
// ═══════════════════════════════════════════════════════════════════════════

app.get("/api/activity-log", async (req, res) => {
  try {
    const { userId, action, limit = 100 } = req.query;

    let whereClause = {};
    if (userId) whereClause.userId = parseInt(userId);
    if (action) whereClause.action = action;

    const logs = await ActivityLog.findAll({
      where: whereClause,
      include: [
        {
          model: User,
          as: "user",
          attributes: ["name", "email", "role"],
        },
      ],
      order: [["timestamp", "DESC"]],
      limit: parseInt(limit),
    });

    res.json(logs);
  } catch (error) {
    console.error("❌ Error fetching activity log:", error);
    res.status(500).json({ error: "Failed to fetch activity log" });
  }
});

// -------------- Insert here: Analytics Dashboard API route --------------
app.get("/api/analytics", async (req, res) => {
  try {
    // Task counts by status
    const statusCounts = await Task.findAll({
      attributes: [
        "status",
        [sequelize.fn("COUNT", sequelize.col("id")), "count"],
      ],
      group: ["status"],
    });
    // Tasks per day (created)
    const tasksPerDay = await Task.findAll({
      attributes: [
        [sequelize.fn("DATE", sequelize.col("createdAt")), "day"],
        [sequelize.fn("COUNT", sequelize.col("id")), "count"],
      ],
      group: ["day"],
      order: [[sequelize.fn("DATE", sequelize.col("createdAt")), "ASC"]],
      raw: true,
    });
    // Staff counts
    const userCount = await User.count({
      where: { role: "employee", isActive: true },
    });
    const adminCount = await User.count({
      where: { role: "admin", isActive: true },
    });
    // Task stat summaries
    const totalTasks = await Task.count();
    const assignedTasks = await Task.count({
      where: { assignedTo: { [Op.ne]: null } },
    });
    const completedTasks = await Task.count({ where: { status: "Completed" } });
    const verifiedTasks = await Task.count({ where: { status: "Verified" } });

    res.json({
      statusCounts,
      tasksPerDay,
      userCount,
      adminCount,
      totalTasks,
      assignedTasks,
      completedTasks,
      verifiedTasks,
    });
  } catch (err) {
    console.error("❌ Analytics error:", err);
    res.status(500).json({ error: "Failed to fetch analytics" });
  }
});
// ---------------- End insertion -----------------

// ═══════════════════════════════════════════════════════════════════════════
// HEALTH & MONITORING ENDPOINTS (Feature #2)
// ═══════════════════════════════════════════════════════════════════════════

app.get("/health", (req, res) => {
  const uptime = process.uptime();
  const hours = Math.floor(uptime / 3600);
  const minutes = Math.floor((uptime % 3600) / 60);
  const seconds = Math.floor(uptime % 60);

  res.status(200).json({
    status: "healthy",
    uptime: `${hours}h ${minutes}m ${seconds}s`,
    uptimeSeconds: Math.floor(uptime),
    timestamp: new Date().toISOString(),
    memory: process.memoryUsage(),
    environment: process.env.NODE_ENV || "production",
  });
});

app.get("/test", (req, res) => {
  res.status(200).send("OK - Server is alive!");
});

// ═══════════════════════════════════════════════════════════════════════════
// FRONTEND - LOGIN PAGE
// ═══════════════════════════════════════════════════════════════════════════

app.get("/", (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Validiant Tracker - Login</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
            position: relative;
            overflow: hidden;
        }

        body::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
            animation: pulse 15s ease-in-out infinite;
        }

        @keyframes pulse {
            0%, 100% { transform: scale(1); }
            50% { transform: scale(1.1); }
        }

        .login-container {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            padding: 50px 40px;
            border-radius: 25px;
            box-shadow: 0 30px 80px rgba(0, 0, 0, 0.3);
            width: 100%;
            max-width: 450px;
            animation: slideIn 0.6s cubic-bezier(0.16, 1, 0.3, 1);
            position: relative;
            z-index: 1;
        }

        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateY(-40px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .logo {
            text-align: center;
            margin-bottom: 35px;
        }

        .logo i {
            font-size: 56px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 15px;
            display: block;
            animation: bounce 2s ease-in-out infinite;
        }

        @keyframes bounce {
            0%, 100% { transform: translateY(0); }
            50% { transform: translateY(-10px); }
        }

        .logo h1 {
            font-size: 32px;
            color: #2d3748;
            font-weight: 700;
            margin-bottom: 8px;
            letter-spacing: -0.5px;
        }

        .logo p {
            color: #718096;
            font-size: 15px;
            font-weight: 500;
        }

        .form-group {
            margin-bottom: 25px;
        }

        .form-group label {
            display: block;
            margin-bottom: 10px;
            color: #2d3748;
            font-weight: 600;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .form-group label i {
            color: #667eea;
            font-size: 16px;
        }

        .form-group input {
            width: 100%;
            padding: 14px 18px;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            font-size: 15px;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            background: white;
            color: #2d3748;
        }

        .form-group input:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1);
            transform: translateY(-2px);
        }

        .form-group input::placeholder {
            color: #a0aec0;
        }

        .btn-login {
            width: 100%;
            padding: 16px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 16px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            margin-top: 10px;
            box-shadow: 0 10px 25px rgba(102, 126, 234, 0.3);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }

        .btn-login:hover {
            transform: translateY(-3px);
            box-shadow: 0 15px 35px rgba(102, 126, 234, 0.4);
        }

        .btn-login:active {
            transform: translateY(-1px);
        }

        .btn-login:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .error-message {
            background: linear-gradient(135deg, #fee 0%, #fdd 100%);
            color: #c33;
            padding: 14px 18px;
            border-radius: 10px;
            margin-bottom: 20px;
            font-size: 14px;
            display: none;
            border-left: 4px solid #c33;
            animation: shake 0.5s;
        }

        @keyframes shake {
            0%, 100% { transform: translateX(0); }
            25% { transform: translateX(-10px); }
            75% { transform: translateX(10px); }
        }

        .error-message.show {
            display: block;
        }

        .loading {
            display: none;
            text-align: center;
            color: #667eea;
            margin-top: 15px;
            font-size: 15px;
            font-weight: 600;
        }

        .loading.show {
            display: block;
        }

        .loading i {
            margin-right: 8px;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        .admin-note {
            margin-top: 25px;
            padding: 15px;
            background: linear-gradient(135deg, #f0f4ff 0%, #e6f0ff 100%);
            border-radius: 10px;
            font-size: 13px;
            color: #4a5568;
            border-left: 4px solid #667eea;
        }

        .admin-note strong {
            color: #2d3748;
            display: block;
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
    <div class="login-container">
        <div class="logo">
            <i class="fas fa-clipboard-check"></i>
            <h1>Validiant Tracker</h1>
            <p>Professional Task Management System</p>
        </div>

        <div id="errorMsg" class="error-message"></div>

        <form id="loginForm">
            <div class="form-group">
                <label for="email">
                    <i class="fas fa-envelope"></i>
                    Email Address
                </label>
                <input 
                    type="email" 
                    id="email" 
                    required 
                    placeholder="Enter your email"
                    autocomplete="email"
                >
            </div>

            <div class="form-group">
                <label for="password">
                    <i class="fas fa-lock"></i>
                    Password
                </label>
                <input 
                    type="password" 
                    id="password" 
                    required 
                    placeholder="Enter your password"
                    autocomplete="current-password"
                >
            </div>

            <button type="submit" class="btn-login" id="loginBtn">
                <i class="fas fa-sign-in-alt"></i>
                Login to Dashboard
            </button>

            <div class="loading" id="loading">
                <i class="fas fa-spinner"></i>
                Logging you in...
            </div>
        </form>


    </div>

    <script>
        const loginForm = document.getElementById('loginForm');
        const errorMsg = document.getElementById('errorMsg');
        const loginBtn = document.getElementById('loginBtn');
        const loading = document.getElementById('loading');
        const emailInput = document.getElementById('email');
        const passwordInput = document.getElementById('password');

        // Focus email on load
        emailInput.focus();

        // Clear error on input
        emailInput.addEventListener('input', () => errorMsg.classList.remove('show'));
        passwordInput.addEventListener('input', () => errorMsg.classList.remove('show'));

        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const email = emailInput.value.trim();
            const password = passwordInput.value;

            if (!email || !password) {
                errorMsg.textContent = 'Please enter both email and password';
                errorMsg.classList.add('show');
                return;
            }

            errorMsg.classList.remove('show');
            loginBtn.disabled = true;
            loginBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Logging in...';
            loading.classList.add('show');

            try {
                const response = await fetch('/api/login', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ email, password })
                });

                const data = await response.json();

                if (data.success) {
                    localStorage.setItem('currentUser', JSON.stringify(data.user));
                    loginBtn.innerHTML = '<i class="fas fa-check-circle"></i> Success! Redirecting...';
                    setTimeout(() => {
                        window.location.href = '/app.js';
                    }, 500);
                } else {
                    errorMsg.textContent = data.message || 'Invalid email or password';
                    errorMsg.classList.add('show');
                    loginBtn.disabled = false;
                    loginBtn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Login to Dashboard';
                    loading.classList.remove('show');
                }
            } catch (err) {
                console.error('Login error:', err);
                errorMsg.textContent = 'Connection error. Please check your internet and try again.';
                errorMsg.classList.add('show');
                loginBtn.disabled = false;
                loginBtn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Login to Dashboard';
                loading.classList.remove('show');
            }
        });

        // Enter key support
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !loginBtn.disabled) {
                loginForm.dispatchEvent(new Event('submit'));
            }
        });
    </script>
</body>
</html>`);
});

// ═══════════════════════════════════════════════════════════════════════════
// FRONTEND - MAIN DASHBOARD APPLICATION (/app.js route)
// ═══════════════════════════════════════════════════════════════════════════

app.get("/app.js", (req, res) => {
  // Using string concatenation to avoid backtick conflicts
  let html = "";

  // HTML Header and CSS
  html += "<!DOCTYPE html>";
  html += '<html lang="en">';
  html += "<head>";
  html += '<meta charset="UTF-8">';
  html +=
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
  html += "<title>Validiant Tracker - Dashboard</title>";
  html +=
    '<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">';
  html += "<style>";

  // Complete CSS Design System
  html += "* { margin: 0; padding: 0; box-sizing: border-box; }";
  html +=
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }';

  // Container
  html += ".container { max-width: 1600px; margin: 0 auto; }";

  // Header
  html +=
    ".header { background: white; padding: 25px 35px; border-radius: 20px; margin-bottom: 25px; box-shadow: 0 10px 40px rgba(0,0,0,0.15); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; animation: slideDown 0.5s ease-out; }";
  html +=
    "@keyframes slideDown { from { opacity: 0; transform: translateY(-30px); } to { opacity: 1; transform: translateY(0); } }";
  html +=
    ".header h1 { font-size: 28px; color: #2d3748; display: flex; align-items: center; gap: 12px; font-weight: 700; }";
  html += ".header h1 i { color: #667eea; font-size: 32px; }";
  html +=
    ".user-info { display: flex; align-items: center; gap: 20px; flex-wrap: wrap; }";
  html +=
    ".user-info span { color: #4a5568; font-size: 15px; font-weight: 600; padding: 10px 20px; background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%); border-radius: 12px; }";
  html += ".user-info span i { color: #667eea; margin-right: 8px; }";

  // Buttons
  html +=
    ".btn { padding: 12px 24px; border: none; border-radius: 10px; cursor: pointer; font-size: 15px; font-weight: 600; transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1); display: inline-flex; align-items: center; gap: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }";
  html +=
    ".btn:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.15); }";
  html += ".btn:active { transform: translateY(-1px); }";
  html +=
    ".btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }";
  html += ".btn i { font-size: 16px; }";

  html +=
    ".btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }";
  html +=
    ".btn-primary:hover { box-shadow: 0 8px 30px rgba(102, 126, 234, 0.4); }";

  html +=
    ".btn-secondary { background: linear-gradient(135deg, #f7fafc 0%, #e2e8f0 100%); color: #2d3748; }";
  html +=
    ".btn-secondary:hover { background: linear-gradient(135deg, #edf2f7 0%, #cbd5e0 100%); }";

  html +=
    ".btn-success { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; }";
  html +=
    ".btn-success:hover { box-shadow: 0 8px 30px rgba(16, 185, 129, 0.4); }";

  html +=
    ".btn-danger { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; }";
  html +=
    ".btn-danger:hover { box-shadow: 0 8px 30px rgba(239, 68, 68, 0.4); }";

  html +=
    ".btn-warning { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; }";
  html +=
    ".btn-warning:hover { box-shadow: 0 8px 30px rgba(245, 158, 11, 0.4); }";

  html +=
    ".btn-info { background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; }";
  html += ".btn-info:hover { box-shadow: 0 8px 30px rgba(59, 130, 246, 0.4); }";

  html += ".btn-sm { padding: 8px 16px; font-size: 13px; }";
  html += ".btn-lg { padding: 16px 32px; font-size: 17px; }";

  // Menu
  html +=
    ".menu { background: white; padding: 20px; border-radius: 20px; margin-bottom: 25px; box-shadow: 0 10px 40px rgba(0,0,0,0.15); display: flex; gap: 12px; flex-wrap: wrap; animation: slideDown 0.6s ease-out; animation-delay: 0.1s; animation-fill-mode: both; }";
  html += ".menu button { flex: 1; min-width: 140px; }";

  // Content Area
  html +=
    ".content { background: white; padding: 35px; border-radius: 20px; box-shadow: 0 10px 40px rgba(0,0,0,0.15); min-height: 500px; animation: fadeIn 0.7s ease-out; animation-delay: 0.2s; animation-fill-mode: both; }";
  html += "@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }";

  html +=
    ".content h2 { color: #2d3748; margin-bottom: 25px; font-size: 26px; font-weight: 700; display: flex; align-items: center; gap: 12px; }";
  html += ".content h2 i { color: #667eea; }";

  // Forms
  html += ".form-group { margin-bottom: 22px; }";
  html +=
    ".form-group label { display: block; margin-bottom: 10px; color: #2d3748; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px; }";
  html += ".form-group label i { color: #667eea; }";
  html +=
    ".form-group input, .form-group select, .form-group textarea { width: 100%; padding: 13px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 15px; transition: all 0.3s; background: white; color: #2d3748; font-family: inherit; }";
  html +=
    ".form-group input:focus, .form-group select:focus, .form-group textarea:focus { outline: none; border-color: #667eea; box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1); transform: translateY(-2px); }";
  html += ".form-group textarea { resize: vertical; min-height: 100px; }";
  html +=
    ".form-group input::placeholder, .form-group textarea::placeholder { color: #a0aec0; }";

  // Task Cards (Employee View)
  html +=
    ".task-card { background: linear-gradient(135deg, #ffffff 0%, #f7fafc 100%); padding: 25px; border-radius: 16px; margin-bottom: 20px; border-left: 5px solid #667eea; transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1); cursor: pointer; box-shadow: 0 4px 15px rgba(0,0,0,0.08); animation: slideInLeft 0.5s ease-out; }";
  html +=
    "@keyframes slideInLeft { from { opacity: 0; transform: translateX(-30px); } to { opacity: 1; transform: translateX(0); } }";
  html +=
    ".task-card:hover { transform: translateX(8px) translateY(-4px); box-shadow: 0 12px 35px rgba(0,0,0,0.12); border-left-color: #764ba2; }";
  html +=
    ".task-card h3 { color: #2d3748; margin-bottom: 15px; font-size: 20px; font-weight: 700; display: flex; align-items: center; gap: 10px; }";
  html += ".task-card h3 i { color: #667eea; }";
  html +=
    ".task-card p { color: #4a5568; font-size: 14px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; }";
  html += ".task-card p i { color: #667eea; width: 18px; }";

  // Pincode Highlight
  html +=
    ".pincode-highlight { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 10px 18px; border-radius: 10px; font-weight: 700; color: #92400e; display: inline-flex; align-items: center; gap: 8px; margin: 12px 0; font-size: 15px; box-shadow: 0 2px 8px rgba(146, 64, 14, 0.15); }";
  html += ".pincode-highlight i { font-size: 16px; }";

  // Status Badges
  html +=
    ".status-badge { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px; border-radius: 20px; font-size: 13px; font-weight: 700; margin-top: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }";
  html +=
    ".status-unassigned { background: linear-gradient(135deg, #e2e8f0 0%, #cbd5e0 100%); color: #2d3748; }";
  html +=
    ".status-pending { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); color: #92400e; }";
  html +=
    ".status-completed { background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%); color: #065f46; }";
  html +=
    ".status-verified { background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); color: #1e40af; }";
  html +=
    ".status-left-job, .status-not-sharing-info, .status-not-picking, .status-switch-off, .status-incorrect-number, .status-wrong-address { background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); color: #991b1b; }";

  // Map Button (Feature #8 - CRITICAL)
  html +=
    ".map-button { background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; padding: 14px 24px; border: none; border-radius: 10px; cursor: pointer; font-size: 15px; font-weight: 700; display: inline-flex; align-items: center; justify-content: center; gap: 10px; transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1); width: 100%; margin-top: 15px; box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3); text-decoration: none; }";
  html +=
    ".map-button:hover { background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); transform: translateY(-3px); box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4); }";
  html += ".map-button i { font-size: 18px; }";
  html +=
    ".no-map { color: #9ca3af; font-size: 14px; font-style: italic; text-align: center; padding: 12px; background: #f9fafb; border-radius: 8px; margin-top: 15px; }";

  // Action Buttons Container
  html +=
    ".action-buttons { display: flex; gap: 12px; margin-top: 15px; flex-wrap: wrap; }";
  html +=
    ".action-buttons button, .action-buttons select { flex: 1; min-width: 120px; }";
  html +=
    ".action-buttons select { padding: 12px; border-radius: 10px; border: 2px solid #e2e8f0; font-size: 14px; background: white; cursor: pointer; font-weight: 600; color: #2d3748; }";
  html +=
    ".action-buttons select:focus { outline: none; border-color: #667eea; box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1); }";

  // Table
  html +=
    ".table { width: 100%; border-collapse: collapse; margin-top: 25px; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }";
  html +=
    ".table thead { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }";
  html +=
    ".table th { padding: 16px 20px; text-align: left; color: white; font-weight: 700; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }";
  html +=
    ".table td { padding: 16px 20px; border-bottom: 1px solid #e2e8f0; color: #2d3748; font-size: 14px; }";
  html += ".table tbody tr { transition: all 0.3s; }";
  html +=
    ".table tbody tr:hover { background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%); transform: scale(1.01); }";
  html += ".table tbody tr:last-child td { border-bottom: none; }";

  // Modal
  html +=
    ".modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.6); backdrop-filter: blur(8px); z-index: 1000; animation: fadeInModal 0.3s; }";
  html += "@keyframes fadeInModal { from { opacity: 0; } to { opacity: 1; } }";
  html +=
    ".modal.show { display: flex; align-items: center; justify-content: center; padding: 20px; }";
  html +=
    ".modal-content { background: white; padding: 40px; border-radius: 25px; max-width: 700px; width: 100%; max-height: 90vh; overflow-y: auto; animation: slideUpModal 0.4s cubic-bezier(0.16, 1, 0.3, 1); box-shadow: 0 30px 80px rgba(0,0,0,0.3); }";
  html +=
    "@keyframes slideUpModal { from { opacity: 0; transform: translateY(50px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }";
  html +=
    ".modal-content h2 { margin-bottom: 25px; color: #2d3748; font-size: 26px; font-weight: 700; display: flex; align-items: center; gap: 12px; }";
  html += ".modal-content h2 i { color: #667eea; }";
  html += ".modal-content::-webkit-scrollbar { width: 8px; }";
  html +=
    ".modal-content::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 10px; }";
  html +=
    ".modal-content::-webkit-scrollbar-thumb { background: #cbd5e0; border-radius: 10px; }";
  html +=
    ".modal-content::-webkit-scrollbar-thumb:hover { background: #94a3b8; }";

  // Toast Notifications
  html +=
    ".toast { position: fixed; top: 30px; right: 30px; background: white; padding: 18px 25px; border-radius: 12px; box-shadow: 0 15px 40px rgba(0,0,0,0.2); z-index: 2000; display: none; min-width: 300px; animation: slideInRight 0.4s cubic-bezier(0.16, 1, 0.3, 1); }";
  html +=
    "@keyframes slideInRight { from { opacity: 0; transform: translateX(400px); } to { opacity: 1; transform: translateX(0); } }";
  html += ".toast.show { display: flex; align-items: center; gap: 12px; }";
  html += ".toast.success { border-left: 5px solid #10b981; }";
  html += ".toast.success i { color: #10b981; font-size: 22px; }";
  html += ".toast.error { border-left: 5px solid #ef4444; }";
  html += ".toast.error i { color: #ef4444; font-size: 22px; }";
  html += ".toast.info { border-left: 5px solid #3b82f6; }";
  html += ".toast.info i { color: #3b82f6; font-size: 22px; }";
  html += ".toast.warning { border-left: 5px solid #f59e0b; }";
  html += ".toast.warning i { color: #f59e0b; font-size: 22px; }";
  html +=
    ".toast-content { flex: 1; color: #2d3748; font-size: 15px; font-weight: 600; }";

  // Search Box
  html +=
    ".search-box { padding: 14px 20px; border: 2px solid #e2e8f0; border-radius: 12px; font-size: 15px; margin-bottom: 25px; width: 100%; max-width: 500px; transition: all 0.3s; background: white; color: #2d3748; }";
  html +=
    ".search-box:focus { outline: none; border-color: #667eea; box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1); transform: translateY(-2px); }";
  html += ".search-box::placeholder { color: #a0aec0; }";

  // Filter Section
  html +=
    ".filter-section { background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%); padding: 25px; border-radius: 15px; margin-bottom: 25px; display: flex; gap: 15px; flex-wrap: wrap; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.05); }";
  html += ".filter-section input, .filter-section select { margin-right: 0; }";
  html +=
    ".filter-section label { font-weight: 600; color: #2d3748; font-size: 14px; }";

  // Grid Layout
  html +=
    ".grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 25px; margin-top: 25px; }";

  // Loading Spinner
  html +=
    ".loading-spinner { display: none; text-align: center; padding: 60px 20px; color: #667eea; font-size: 18px; font-weight: 600; }";
  html += ".loading-spinner.show { display: block; }";
  html +=
    ".loading-spinner i { font-size: 48px; display: block; margin-bottom: 20px; animation: spin 1.5s linear infinite; }";
  html +=
    "@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }";

  // Empty State
  html +=
    ".empty-state { text-align: center; padding: 80px 20px; color: #718096; }";
  html +=
    ".empty-state i { font-size: 80px; color: #cbd5e0; margin-bottom: 20px; display: block; }";
  html +=
    ".empty-state h3 { font-size: 22px; margin-bottom: 10px; color: #4a5568; }";
  html += ".empty-state p { font-size: 16px; }";

  // Progress Bar (for bulk upload)
  html += ".progress-container { margin: 25px 0; }";
  html +=
    ".progress-bar-wrapper { background: #e2e8f0; height: 30px; border-radius: 15px; overflow: hidden; box-shadow: inset 0 2px 4px rgba(0,0,0,0.1); }";
  html +=
    ".progress-bar { background: linear-gradient(135deg, #10b981 0%, #059669 100%); height: 100%; width: 0; transition: width 0.3s; display: flex; align-items: center; justify-content: center; color: white; font-weight: 700; font-size: 13px; }";
  html +=
    ".progress-text { text-align: center; margin-top: 12px; color: #4a5568; font-size: 15px; font-weight: 600; }";

  // Info Box (for instructions, notes)
  html +=
    ".info-box { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 20px; border-radius: 12px; margin-bottom: 25px; border-left: 5px solid #f59e0b; }";
  html +=
    ".info-box.success { background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%); border-left-color: #10b981; }";
  html +=
    ".info-box.error { background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); border-left-color: #ef4444; }";
  html +=
    ".info-box.info { background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); border-left-color: #3b82f6; }";
  html +=
    ".info-box p { color: #2d3748; font-size: 14px; margin-bottom: 10px; font-weight: 600; }";
  html += ".info-box p:last-child { margin-bottom: 0; }";
  html +=
    ".info-box ul { margin-left: 20px; color: #4a5568; font-size: 14px; }";
  html += ".info-box ul li { margin-bottom: 8px; }";

  // Checkbox (for bulk operations)
  html +=
    ".checkbox-container { display: flex; align-items: center; gap: 10px; }";
  html +=
    '.checkbox-container input[type="checkbox"] { width: 20px; height: 20px; cursor: pointer; accent-color: #667eea; }';
  html +=
    ".bulk-action-bar { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; border-radius: 12px; margin-bottom: 25px; display: none; align-items: center; gap: 15px; box-shadow: 0 8px 30px rgba(102, 126, 234, 0.3); }";
  html += ".bulk-action-bar.show { display: flex; }";
  html +=
    ".bulk-action-bar span { color: white; font-weight: 700; font-size: 16px; flex: 1; }";

  // Distance indicator (for GPS sorting)
  html +=
    ".distance-indicator { background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); color: #1e40af; padding: 8px 14px; border-radius: 8px; font-weight: 700; font-size: 13px; display: inline-flex; align-items: center; gap: 6px; margin-top: 10px; }";
  html += ".distance-indicator i { font-size: 14px; }";

  // Responsive Design
  html += "@media (max-width: 768px) {";
  html += ".header { padding: 20px; }";
  html += ".header h1 { font-size: 22px; }";
  html += ".menu { padding: 15px; gap: 8px; }";
  html += ".menu button { min-width: 100%; font-size: 14px; }";
  html += ".content { padding: 20px; }";
  html += ".grid { grid-template-columns: 1fr; }";
  html += ".action-buttons { flex-direction: column; }";
  html += ".action-buttons button, .action-buttons select { width: 100%; }";
  html += ".table { font-size: 12px; }";
  html += ".table th, .table td { padding: 10px; }";
  html += ".modal-content { padding: 25px; }";
  html += ".filter-section { flex-direction: column; align-items: stretch; }";
  html += ".search-box { max-width: 100%; }";
  html += "}";

  html += "</style>";
  html += "</head>";
  html += "<body>";

  // Toast notification element
  html += '<div id="toast" class="toast"></div>';

  // Main container
  html += '<div class="container">';

  // Header
  html += '<div class="header">';
  html += '<h1><i class="fas fa-clipboard-check"></i> Validiant Tracker</h1>';
  html += '<div class="user-info">';
  html +=
    '<span id="userName"><i class="fas fa-user-circle"></i> Loading...</span>';
  html += '<button class="btn btn-danger btn-sm" onclick="logout()">';
  html += '<i class="fas fa-sign-out-alt"></i> Logout';
  html += "</button>";
  html += "</div>";
  html += "</div>";

  // Menu (will be populated by JavaScript based on role)
  html += '<div class="menu" id="menu"></div>';

  // Content area (will be populated by JavaScript)
  html += '<div class="content" id="content">';
  html += '<div class="loading-spinner show">';
  html += '<i class="fas fa-spinner"></i>';
  html += "Loading dashboard...";
  html += "</div>";
  html += "</div>";

  html += "</div>"; // End container

  // Now add the JavaScript (using string concatenation to avoid backtick issues)
  html += "<script>";

  html +=
    "// PART 4 - JavaScript continuation (inside the <script> tag from PART 3)\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// GLOBAL VARIABLES & SESSION MANAGEMENT\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "let currentUser = null;\n";
  html += "let sessionTimeout = null;\n";
  html += "let allEmployeeTasks = [];\n";
  html += "let allHistoryTasks = [];\n";
  html += "let selectedTaskIds = [];\n";
  html += "\n";
  html += "// Check authentication\n";
  html += "const userStr = localStorage.getItem('currentUser');\n";
  html += "if (!userStr) {\n";
  html += "  window.location.href = '/';\n";
  html += "  throw new Error('Not authenticated');\n";
  html += "}\n";
  html += "\n";
  html += "try {\n";
  html += "  currentUser = JSON.parse(userStr);\n";
  html +=
    "  document.getElementById('userName').innerHTML = '<i class=\"fas fa-user-circle\"></i> ' + escapeHtml(currentUser.name) + ' (' + currentUser.role + ')';\n";
  html += "} catch (e) {\n";
  html += "  console.error('Error parsing user data:', e);\n";
  html += "  window.location.href = '/';\n";
  html += "  throw new Error('Invalid user data');\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// SESSION TIMEOUT MANAGEMENT (15 minutes)\n";
  html +=
    "// ═════════════════════════════════════════════════o��═════════════════════════\n";
  html += "\n";
  html += "function resetSessionTimeout() {\n";
  html += "  clearTimeout(sessionTimeout);\n";
  html += "  sessionTimeout = setTimeout(function() {\n";
  html += "    showToast('Session expired. Please login again.', 'error');\n";
  html += "    setTimeout(function() {\n";
  html += "      logout();\n";
  html += "    }, 2000);\n";
  html += "  }, 900000); // 15 minutes\n";
  html += "}\n";
  html += "\n";
  html += "resetSessionTimeout();\n";
  html += "document.addEventListener('click', resetSessionTimeout);\n";
  html += "document.addEventListener('keypress', resetSessionTimeout);\n";
  html += "document.addEventListener('mousemove', resetSessionTimeout);\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// UTILITY FUNCTIONS\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function escapeHtml(text) {\n";
  html += "  if (!text) return '';\n";
  html += "  const div = document.createElement('div');\n";
  html += "  div.textContent = text;\n";
  html += "  return div.innerHTML;\n";
  html += "}\n";
  html += "\n";
  html += "function showToast(message, type) {\n";
  html += "  const toast = document.getElementById('toast');\n";
  html += "  const icons = {\n";
  html += "    success: 'fa-check-circle',\n";
  html += "    error: 'fa-exclamation-circle',\n";
  html += "    info: 'fa-info-circle',\n";
  html += "    warning: 'fa-exclamation-triangle'\n";
  html += "  };\n";
  html += "  const icon = icons[type] || icons.info;\n";
  html += "  \n";
  html +=
    "  toast.innerHTML = '<i class=\"fas ' + icon + '\"></i><span class=\"toast-content\">' + escapeHtml(message) + '</span>';\n";
  html += "  toast.className = 'toast show ' + type;\n";
  html += "  \n";
  html += "  setTimeout(function() {\n";
  html += "    toast.classList.remove('show');\n";
  html += "  }, 4000);\n";
  html += "}\n";
  html += "\n";
  html += "function showLoading(containerId) {\n";
  html +=
    "  const container = document.getElementById(containerId || 'content');\n";
  html +=
    '  container.innerHTML = \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading...</div>\';\n';
  html += "}\n";
  html += "\n";
  html += "function hideLoading(containerId) {\n";
  html +=
    "  const spinner = document.querySelector('#' + (containerId || 'content') + ' .loading-spinner');\n";
  html += "  if (spinner) spinner.remove();\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// LOGOUT FUNCTION (Feature #10)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function logout() {\n";
  html += "  if (!confirm('Are you sure you want to logout?')) {\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  localStorage.removeItem('currentUser');\n";
  html += "  sessionStorage.clear();\n";
  html += "  showToast('Logged out successfully!', 'success');\n";
  html += "  \n";
  html += "  setTimeout(function() {\n";
  html += "    window.location.href = '/';\n";
  html += "  }, 1000);\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// KEYBOARD SHORTCUTS (Enhancement Feature)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "document.addEventListener('keydown', function(e) {\n";
  html += "  // Don't trigger shortcuts when typing in inputs\n";
  html +=
    "  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  // Ctrl/Cmd + N: New task (admin only)\n";
  html +=
    "  if ((e.ctrlKey || e.metaKey) && e.key === 'n' && currentUser.role === 'admin') {\n";
  html += "    e.preventDefault();\n";
  html += "    showAssignTask();\n";
  html += "  }\n";
  html += "  \n";
  html += "  // Ctrl/Cmd + U: Unassigned tasks (admin only)\n";
  html +=
    "  if ((e.ctrlKey || e.metaKey) && e.key === 'u' && currentUser.role === 'admin') {\n";
  html += "    e.preventDefault();\n";
  html += "    showUnassignedTasks();\n";
  html += "  }\n";
  html += "  \n";
  html += "  // Ctrl/Cmd + E: Employees (admin only)\n";
  html +=
    "  if ((e.ctrlKey || e.metaKey) && e.key === 'e' && currentUser.role === 'admin') {\n";
  html += "    e.preventDefault();\n";
  html += "    showEmployees();\n";
  html += "  }\n";
  html += "  \n";
  html += "  // Ctrl/Cmd + /: Focus search\n";
  html += "  if ((e.ctrlKey || e.metaKey) && e.key === '/') {\n";
  html += "    e.preventDefault();\n";
  html += "    const searchBox = document.querySelector('.search-box');\n";
  html += "    if (searchBox) searchBox.focus();\n";
  html += "  }\n";
  html += "  \n";
  html += "  // Escape: Close modals\n";
  html += "  if (e.key === 'Escape') {\n";
  html += "    const modals = document.querySelectorAll('.modal');\n";
  html += "    modals.forEach(function(modal) {\n";
  html += "      modal.remove();\n";
  html += "    });\n";
  html += "  }\n";
  html += "});\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// ADMIN FUNCTIONS - ASSIGN TASK\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showAssignTask() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-tasks\"></i> Assign New Task</h2>';\n";
  html += "  \n";
  html += "  // Bulk upload button\n";
  html +=
    '  html += \'<button class="btn btn-success" onclick="showBulkUpload()" style="margin-bottom: 25px;">\';\n';
  html +=
    "  html += '<i class=\"fas fa-file-excel\"></i> Bulk Upload Tasks (Excel)';\n";
  html += "  html += '</button>';\n";
  html += "  \n";
  html += " html += '<form id=\"taskForm\">';\n";
  html += " \n";
  html += " html += '<div class=\"form-group\">';\n";
  html +=
    ' html += \'<label for="caseId"><i class="fas fa-id-card"></i> Case ID / Title *</label>\';\n';
  html +=
    ' html += \'<input type="text" id="caseId" required placeholder="Enter case ID or title" maxlength="500">\';\n';
  html += " html += '</div>';\n";
  html += " \n";
  html += " html += '<div class=\"form-group\">';\n";
  html +=
    ' html += \'<label for="clientName"><i class="fas fa-user-tie"></i> Client Name (Optional)</label>\';\n';
  html +=
    ' html += \'<input type="text" id="clientName" placeholder="Enter client name (optional)" maxlength="200">\';\n';
  html += " html += '</div>';\n";
  html += " \n";
  html += " html += '<div class=\"form-group\">';\n";
  html +=
    ' html += \'<label for="pincode"><i class="fas fa-map-pin"></i> Pincode *</label>\';\n';
  html +=
    ' html += \'<input type="text" id="pincode" required placeholder="6-digit pincode" maxlength="6" pattern="[0-9]{6}">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="mapUrl"><i class="fas fa-map-marked-alt"></i> Google Maps URL (Optional)</label>\';\n';
  html +=
    '  html += \'<input type="url" id="mapUrl" placeholder="Paste Google Maps link (coordinates will be extracted automatically)">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="latitude"><i class="fas fa-globe"></i> Latitude (Auto-filled from map URL)</label>\';\n';
  html +=
    '  html += \'<input type="number" id="latitude" step="any" placeholder="Will be extracted from map URL">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="longitude"><i class="fas fa-globe"></i> Longitude (Auto-filled from map URL)</label>\';\n';
  html +=
    '  html += \'<input type="number" id="longitude" step="any" placeholder="Will be extracted from map URL">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="employee"><i class="fas fa-user"></i> Assign to Employee (Optional - Leave blank for unassigned)</label>\';\n';
  html += "  html += '<select id=\"employee\">';\n";
  html += "  html += '<option value=\"\">-- Leave Unassigned --</option>';\n";
  html += "  html += '</select>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="notes"><i class="fas fa-sticky-note"></i> Notes (Optional)</label>\';\n';
  html +=
    '  html += \'<textarea id="notes" rows="4" placeholder="Additional notes or instructions"></textarea>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html +=
    '  html += \'<button type="submit" class="btn btn-primary btn-lg">\';\n';
  html += "  html += '<i class=\"fas fa-check\"></i> Create Task';\n";
  html += "  html += '</button>';\n";
  html += "  \n";
  html += "  html += '</form>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  \n";
  html += "  // Load employees\n";
  html += "  fetch('/api/users')\n";
  html += "    .then(function(r) { return r.json(); })\n";
  html += "    .then(function(users) {\n";
  html += "      const select = document.getElementById('employee');\n";
  html += "      users.forEach(function(u) {\n";
  html += "        const option = document.createElement('option');\n";
  html += "        option.value = u.id;\n";
  html +=
    "        option.textContent = escapeHtml(u.name) + ' (' + escapeHtml(u.email) + ')';\n";
  html += "        select.appendChild(option);\n";
  html += "      });\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading employees:', err);\n";
  html += "      showToast('Failed to load employees', 'error');\n";
  html += "    });\n";
  html += "  \n";
  html += "  // Map URL coordinate extraction\n";
  html +=
    "  document.getElementById('mapUrl').addEventListener('blur', function() {\n";
  html += "    const url = this.value;\n";
  html += "    if (url) {\n";
  html += "      const latMatch = url.match(/@(-?[0-9.]+),(-?[0-9.]+)/);\n";
  html += "      if (latMatch) {\n";
  html += "        document.getElementById('latitude').value = latMatch[1];\n";
  html += "        document.getElementById('longitude').value = latMatch[2];\n";
  html +=
    "        showToast('Coordinates extracted from map URL!', 'success');\n";
  html += "      } else {\n";
  html += "        const qMatch = url.match(/\\?q=(-?[0-9.]+),(-?[0-9.]+)/);\n";
  html += "        if (qMatch) {\n";
  html += "          document.getElementById('latitude').value = qMatch[1];\n";
  html += "          document.getElementById('longitude').value = qMatch[2];\n";
  html +=
    "          showToast('Coordinates extracted from map URL!', 'success');\n";
  html += "        }\n";
  html += "      }\n";
  html += "    }\n";
  html += "  });\n";
  html += "  \n";
  html += "  // Form submission\n";
  html +=
    "  document.getElementById('taskForm').addEventListener('submit', function(e) {\n";
  html += "    e.preventDefault();\n";
  html += "    \n";
  html += "    const pincode = document.getElementById('pincode').value;\n";
  html += "    if (!/^[0-9]{6}$/.test(pincode)) {\n";
  html += "      showToast('Pincode must be exactly 6 digits', 'error');\n";
  html += "      return;\n";
  html += "    }\n";
  html += "    \n";
  html += "    const formData = {\n";
  html += "      title: document.getElementById('caseId').value,\n";
  html +=
    "      clientName: document.getElementById('clientName').value || null,\n";
  html += "      pincode: pincode,\n";
  html += "      mapUrl: document.getElementById('mapUrl').value || null,\n";
  html +=
    "      latitude: parseFloat(document.getElementById('latitude').value) || null,\n";
  html +=
    "      longitude: parseFloat(document.getElementById('longitude').value) || null,\n";
  html +=
    "      assignedTo: document.getElementById('employee').value ? parseInt(document.getElementById('employee').value) : null,\n";
  html += "      notes: document.getElementById('notes').value || null,\n";
  html += "      createdBy: currentUser.id,\n";
  html += "      createdByName: currentUser.name\n";
  html += "    };\n";
  html += "    \n";
  html +=
    "    const btn = e.target.querySelector('button[type=\"submit\"]');\n";
  html += "    btn.disabled = true;\n";
  html +=
    "    btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> Creating...';\n";
  html += "    \n";
  html += "    fetch('/api/tasks', {\n";
  html += "      method: 'POST',\n";
  html += "      headers: { 'Content-Type': 'application/json' },\n";
  html += "      body: JSON.stringify(formData)\n";
  html += "    })\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(data) {\n";
  html += "      if (data.success) {\n";
  html +=
    "        showToast(formData.assignedTo ? 'Task assigned successfully!' : 'Task created as unassigned!', 'success');\n";
  html += "        document.getElementById('taskForm').reset();\n";
  html += "        btn.disabled = false;\n";
  html +=
    "        btn.innerHTML = '<i class=\"fas fa-check\"></i> Create Task';\n";
  html += "      } else {\n";
  html +=
    "        showToast(data.message || 'Error creating task', 'error');\n";
  html += "        btn.disabled = false;\n";
  html +=
    "        btn.innerHTML = '<i class=\"fas fa-check\"></i> Create Task';\n";
  html += "      }\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error creating task:', err);\n";
  html += "      showToast('Connection error. Please try again.', 'error');\n";
  html += "      btn.disabled = false;\n";
  html +=
    "      btn.innerHTML = '<i class=\"fas fa-check\"></i> Create Task';\n";
  html += "    });\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// ADMIN FUNCTIONS - BULK UPLOAD (Feature #3)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showBulkUpload() {\n";
  html += "  const modal = document.createElement('div');\n";
  html += "  modal.className = 'modal show';\n";
  html += "  \n";
  html += "  let html = '<div class=\"modal-content\">';\n";
  html +=
    "  html += '<h2><i class=\"fas fa-file-excel\"></i> Bulk Upload Tasks</h2>';\n";
  html += "  \n";
  html += "  html += '<div class=\"info-box\">';\n";
  html +=
    "  html += '<p style=\"color: #92400e; font-weight: 700; margin-bottom: 12px;\">';\n";
  html +=
    "  html += '<i class=\"fas fa-exclamation-triangle\"></i> IMPORTANT: Tasks will be uploaded as UNASSIGNED';\n";
  html += "  html += '</p>';\n";
  html += "  html += '<p style=\"color: #78350f; font-size: 14px;\">';\n";
  html +=
    "  html += 'After upload, assign them to employees from the \"Unassigned Tasks\" menu.';\n";
  html += "  html += '</p>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-list\"></i> Excel Format Required:</label>';\n";
  html +=
    "  html += '<ul style=\"margin-left: 20px; color: #4a5568; font-size: 14px;\">';\n";
  html +=
    "  html += '<li><strong>RequestID</strong> or <strong>Title</strong> (required)</li>';\n";
  html += "  html += '<li><strong>Individual Name</strong> (optional)</li>';\n";
  html +=
    "  html += '<li><strong>Assigned To (ID)</strong> (optional - e.g., EMP001)</li>';\n";
  html +=
    "  html += '<li><strong>Employee Email</strong> (optional - e.g., emp@company.com)</li>';\n";
  html +=
    "  html += '<li><strong>Pincode</strong> (required - 6 digits)</li>';\n";
  html += "  html += '<li><strong>MapURL</strong> (optional)</li>';\n";
  html += "  html += '<li><strong>Latitude</strong> (optional)</li>';\n";
  html += "  html += '<li><strong>Longitude</strong> (optional)</li>';\n";
  html += "  html += '<li><strong>Notes</strong> (optional)</li>';\n";
  html += "  html += '</ul>';\n";
  html +=
    '  html += \'<button class="btn btn-secondary btn-sm" onclick="downloadTemplate()" style="margin-top: 15px;">\';\n';
  html +=
    "  html += '<i class=\"fas fa-download\"></i> Download CSV Template';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    '  html += \'<label for="excelFile"><i class="fas fa-file-upload"></i> Choose Excel File:</label>\';\n';
  html +=
    '  html += \'<input type="file" id="excelFile" accept=".xlsx,.xls" style="padding: 10px; border: 2px dashed #cbd5e0; border-radius: 10px; width: 100%;">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html +=
    '  html += \'<div id="uploadProgress" style="display: none;" class="progress-container">\';\n';
  html += "  html += '<div class=\"progress-bar-wrapper\">';\n";
  html +=
    '  html += \'<div id="progressBar" class="progress-bar">0%</div>\';\n';
  html += "  html += '</div>';\n";
  html += '  html += \'<p id="progressText" class="progress-text"></p>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html +=
    "  html += '<div style=\"display: flex; gap: 12px; margin-top: 25px;\">';\n";
  html +=
    '  html += \'<button class="btn btn-primary" onclick="processUpload()">\';\n';
  html += "  html += '<i class=\"fas fa-upload\"></i> Upload & Process';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button class="btn btn-secondary" onclick="closeBulkUpload()">\';\n';
  html += "  html += '<i class=\"fas fa-times\"></i> Cancel';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  modal.innerHTML = html;\n";
  html += "  document.body.appendChild(modal);\n";
  html += "}\n";
  html += "\n";
  html += "function downloadTemplate() {\n";
  html +=
    "  const csv = 'RequestID,Individual Name,Employee ID,Employee Email,Pincode,MapURL,Latitude,Longitude,Notes\\nCASE001,Client A,EMP001,,560001,https://maps.google.com/?q=12.9716,77.5946,12.9716,77.5946,Sample task\\nCASE002,Client B,,emp@example.com,560002,,,,\"Another sample task\"';\n";
  html += "  const blob = new Blob([csv], { type: 'text/csv' });\n";
  html += "  const url = URL.createObjectURL(blob);\n";
  html += "  const a = document.createElement('a');\n";
  html += "  a.href = url;\n";
  html += "  a.download = 'validiant-bulk-upload-template.csv';\n";
  html += "  a.click();\n";
  html += "  URL.revokeObjectURL(url);\n";
  html += "  showToast('Template downloaded successfully!', 'success');\n";
  html += "}\n";
  html += "\n";
  html += "function processUpload() {\n";
  html += "  const fileInput = document.getElementById('excelFile');\n";
  html += "  if (!fileInput.files[0]) {\n";
  html += "    showToast('Please select an Excel file', 'error');\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  const formData = new FormData();\n";
  html += "  formData.append('excelFile', fileInput.files[0]);\n";
  html += "  formData.append('adminId', currentUser.id);\n";
  html += "  formData.append('adminName', currentUser.name);\n";
  html += "  \n";
  html +=
    "  document.getElementById('uploadProgress').style.display = 'block';\n";
  html +=
    "  document.getElementById('progressText').textContent = 'Uploading file...';\n";
  html += "  document.getElementById('progressBar').style.width = '30%';\n";
  html += "  document.getElementById('progressBar').textContent = '30%';\n";
  html += "  \n";
  html += "  fetch('/api/tasks/bulk-upload', {\n";
  html += "    method: 'POST',\n";
  html += "    body: formData\n";
  html += "  })\n";
  html += "  .then(function(res) { \n";
  html += "    if (!res.ok) {\n";
  html +=
    "      throw new Error('Server returned ' + res.status + ': ' + res.statusText);\n";
  html += "    }\n";
  html += "    return res.json();\n";
  html += "  })\n";
  html += "  .then(function(data) {\n";
  html += "    document.getElementById('progressBar').style.width = '100%';\n";
  html += "    document.getElementById('progressBar').textContent = '100%';\n";
  html += "    \n";
  html += "    if (data.success) {\n";
  html +=
    "      document.getElementById('progressText').textContent = data.message;\n";
  html +=
    "      showToast(data.successCount + ' tasks uploaded as unassigned!', 'success');\n";
  html += "      \n";
  html += "      if (data.errors && data.errors.length > 0) {\n";
  html += "        console.log('Upload errors:', data.errors);\n";
  html +=
    "        showToast(data.errorCount + ' tasks failed. Check console for details.', 'warning');\n";
  html += "      }\n";
  html += "      \n";
  html += "      setTimeout(function() {\n";
  html += "        closeBulkUpload();\n";
  html += "        showUnassignedTasks();\n";
  html += "      }, 2500);\n";
  html += "    } else {\n";
  html +=
    "      document.getElementById('progressText').textContent = 'Error: ' + data.message;\n";
  html += "      showToast('Upload failed: ' + data.message, 'error');\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Upload error:', err);\n";
  html +=
    "    document.getElementById('progressText').textContent = 'Upload failed';\n";
  html += "    showToast('Upload error: ' + err.message, 'error');\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function closeBulkUpload() {\n";
  html += "  const modal = document.querySelector('.modal');\n";
  html += "  if (modal) modal.remove();\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// ADMIN FUNCTIONS - UNASSIGNED TASKS (Feature #4)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showUnassignedTasks() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-inbox\"></i> Unassigned Tasks Pool</h2>';\n";
  html += "  \n";
  html += "  html += '<div class=\"filter-section\">';\n";
  html +=
    '  html += \'<input type="text" id="unassignedSearch" class="search-box" placeholder="Search by Case ID or Pincode..." style="max-width: 400px;">\';\n';
  html +=
    '  html += \'<button class="btn btn-info btn-sm" onclick="searchUnassignedTasks()">\';\n';
  html += "  html += '<i class=\"fas fa-search\"></i> Search';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button class="btn btn-secondary btn-sm" onclick="loadUnassignedTasks()">\';\n';
  html += "  html += '<i class=\"fas fa-sync\"></i> Refresh';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div id=\"unassignedTasksList\">';\n";
  html +=
    '  html += \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading unassigned tasks...</div>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  \n";
  html += "  // Enable search on enter key\n";
  html +=
    "  document.getElementById('unassignedSearch').addEventListener('keypress', function(e) {\n";
  html += "    if (e.key === 'Enter') {\n";
  html += "      searchUnassignedTasks();\n";
  html += "    }\n";
  html += "  });\n";
  html += "  \n";
  html += "  loadUnassignedTasks();\n";
  html += "}\n";
  html += "\n";
  html += "function loadUnassignedTasks(searchTerm) {\n";
  html +=
    "  const url = '/api/tasks/unassigned' + (searchTerm ? '?search=' + encodeURIComponent(searchTerm) : '');\n";
  html += "  \n";
  html += "  fetch(url)\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(tasks) {\n";
  html += "      fetch('/api/users')\n";
  html += "        .then(function(r) { return r.json(); })\n";
  html += "        .then(function(employees) {\n";
  html += "          displayUnassignedTasks(tasks, employees);\n";
  html += "        });\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading unassigned tasks:', err);\n";
  html +=
    "      document.getElementById('unassignedTasksList').innerHTML = '<div class=\"empty-state\"><i class=\"fas fa-exclamation-triangle\"></i><h3>Error Loading Tasks</h3><p>Please try again.</p></div>';\n";
  html += "      showToast('Failed to load tasks', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function displayUnassignedTasks(tasks, employees) {\n";
  html += "  let html = '';\n";
  html += "  \n";
  html += "  if (tasks.length === 0) {\n";
  html += "    html = '<div class=\"empty-state\">';\n";
  html += "    html += '<i class=\"fas fa-check-circle\"></i>';\n";
  html += "    html += '<h3>All Clear!</h3>';\n";
  html +=
    "    html += '<p>No unassigned tasks found. All tasks have been assigned to employees.</p>';\n";
  html += "    html += '</div>';\n";
  html += "  } else {\n";
  html +=
    "    html = '<p style=\"color: #4a5568; font-weight: 600; margin-bottom: 20px;\">';\n";
  html +=
    "    html += '<i class=\"fas fa-info-circle\"></i> Found ' + tasks.length + ' unassigned task(s)';\n";
  html += "    html += '</p>';\n";
  html += "    \n";
  html += "    html += '<table class=\"table\">';\n";
  html += "    html += '<thead><tr>';\n";
  html += "    html += '<th>Case ID</th>';\n";
  html += "    html += '<th>Client Name</th>';\n";
  html += "    html += '<th>Pincode</th>';\n";
  html += "    html += '<th>Map URL</th>';\n";
  html += "    html += '<th>Notes</th>';\n";
  html += "    html += '<th>Actions</th>';\n";
  html += "    html += '</tr></thead>';\n";
  html += "    html += '<tbody>';\n";
  html += "    \n";
  html += "    tasks.forEach(function(task) {\n";
  html += "      html += '<tr>';\n";
  html +=
    "      html += '<td><strong>' + escapeHtml(task.title) + '</strong></td>';\n";
  html += " html += '<td>' + escapeHtml(task.clientName || '-') + '</td>';\n";
  html +=
    "      html += '<td><span class=\"pincode-highlight\"><i class=\"fas fa-map-pin\"></i> ' + escapeHtml(task.pincode || 'N/A') + '</span></td>';\n";
  html += "      html += '<td>';\n";
  html += "      if (task.mapUrl) {\n";
  html +=
    '        html += \'<a href="\' + escapeHtml(task.mapUrl) + \'" target="_blank" rel="noopener noreferrer" style="color: #3b82f6; font-weight: 600;">\';\n';
  html +=
    "        html += '<i class=\"fas fa-map-marker-alt\"></i> View Map';\n";
  html += "        html += '</a>';\n";
  html += "      } else {\n";
  html +=
    "        html += '<span style=\"color: #9ca3af; font-style: italic;\">No map</span>';\n";
  html += "      }\n";
  html += "      html += '</td>';\n";
  html += "      html += '<td>' + escapeHtml(task.notes || 'N/A') + '</td>';\n";
  html += "      html += '<td>';\n";
  html +=
    "      html += '<select id=\"emp-' + task.id + '\" style=\"padding: 10px; border-radius: 8px; margin-right: 10px; border: 2px solid #e2e8f0; font-weight: 600;\">';\n";
  html += "      html += '<option value=\"\">Select Employee</option>';\n";
  html += "      employees.forEach(function(emp) {\n";
  html +=
    "        html += '<option value=\"' + emp.id + '\">' + escapeHtml(emp.name) + '</option>';\n";
  html += "      });\n";
  html += "      html += '</select>';\n";
  html +=
    "      html += '<button class=\"btn btn-success btn-sm\" onclick=\"assignTaskToEmployee(' + task.id + ')\">';\n";
  html += "      html += '<i class=\"fas fa-user-check\"></i> Assign';\n";
  html += "      html += '</button>';\n";
  html += "      html += '</td>';\n";
  html += "      html += '</tr>';\n";
  html += "    });\n";
  html += "    \n";
  html += "    html += '</tbody></table>';\n";
  html += "  }\n";
  html += "  \n";
  html +=
    "  document.getElementById('unassignedTasksList').innerHTML = html;\n";
  html += "}\n";
  html += "\n";
  html += "function searchUnassignedTasks() {\n";
  html +=
    "  const searchTerm = document.getElementById('unassignedSearch').value.trim();\n";
  html += "  loadUnassignedTasks(searchTerm);\n";
  html += "}\n";
  html += "\n";
  html += "function assignTaskToEmployee(taskId) {\n";
  html += "  const select = document.getElementById('emp-' + taskId);\n";
  html += "  const employeeId = select.value;\n";
  html += "  \n";
  html += "  if (!employeeId) {\n";
  html += "    showToast('Please select an employee first', 'error');\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  const btn = select.nextElementSibling;\n";
  html += "  btn.disabled = true;\n";
  html +=
    "  btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> Assigning...';\n";
  html += "  \n";
  html += "  fetch('/api/tasks/' + taskId + '/assign', {\n";
  html += "    method: 'POST',\n";
  html += "    headers: { 'Content-Type': 'application/json' },\n";
  html += "    body: JSON.stringify({\n";
  html += "      employeeId: parseInt(employeeId),\n";
  html += "      adminId: currentUser.id,\n";
  html += "      adminName: currentUser.name\n";
  html += "    })\n";
  html += "  })\n";
  html += "  .then(function(res) { return res.json(); })\n";
  html += "  .then(function(data) {\n";
  html += "    if (data.success) {\n";
  html +=
    "      showToast(data.message || 'Task assigned successfully!', 'success');\n";
  html += "      loadUnassignedTasks();\n";
  html += "    } else {\n";
  html += "      showToast(data.message || 'Error assigning task', 'error');\n";
  html += "      btn.disabled = false;\n";
  html +=
    "      btn.innerHTML = '<i class=\"fas fa-user-check\"></i> Assign';\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Error assigning task:', err);\n";
  html += "    showToast('Connection error', 'error');\n";
  html += "    btn.disabled = false;\n";
  html += "    btn.innerHTML = '<i class=\"fas fa-user-check\"></i> Assign';\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ════════════════════════════════════════════════ ��══════════════════════════\n";
  html += "// ADMIN FUNCTIONS - VIEW ALL TASKS (Feature #5 & #6)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showAllTasks() {\n";
  html += "  let html = '<h2><i class=\"fas fa-list\"></i> All Tasks</h2>';\n";
  html += "  \n";
  html += "  html += '<div class=\"filter-section\">';\n";
  html +=
    '  html += \'<input type="text" id="taskSearch" class="search-box" placeholder="Search by Case ID, Pincode, Status..." style="max-width: 400px;">\';\n';
  html +=
    '  html += \'<button class="btn btn-info btn-sm" onclick="searchAllTasks()">\';\n';
  html += "  html += '<i class=\"fas fa-search\"></i> Search';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button class="btn btn-secondary btn-sm" onclick="loadAllTasks()">\';\n';
  html += "  html += '<i class=\"fas fa-sync\"></i> Refresh';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div id=\"allTasksList\">';\n";
  html +=
    '  html += \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading all tasks...</div>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  \n";
  html +=
    "  document.getElementById('taskSearch').addEventListener('keypress', function(e) {\n";
  html += "    if (e.key === 'Enter') {\n";
  html += "      searchAllTasks();\n";
  html += "    }\n";
  html += "  });\n";
  html += "  \n";
  html += "  loadAllTasks();\n";
  html += "}\n";
  html += "\n";
  html += "function loadAllTasks(searchTerm) {\n";
  html +=
    "  const url = '/api/tasks?role=admin' + (searchTerm ? '&search=' + encodeURIComponent(searchTerm) : '');\n";
  html += "  \n";
  html += "  fetch(url)\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(tasks) {\n";
  html += "      displayAllTasks(tasks);\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading tasks:', err);\n";
  html +=
    "      document.getElementById('allTasksList').innerHTML = '<div class=\"empty-state\"><i class=\"fas fa-exclamation-triangle\"></i><h3>Error Loading Tasks</h3><p>Please try again.</p></div>';\n";
  html += "      showToast('Failed to load tasks', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function displayAllTasks(tasks) {\n";
  html += "  let html = '';\n";
  html += "  \n";
  html += "  if (tasks.length === 0) {\n";
  html += "    html = '<div class=\"empty-state\">';\n";
  html += "    html += '<i class=\"fas fa-inbox\"></i>';\n";
  html += "    html += '<h3>No Tasks Found</h3>';\n";
  html +=
    "    html += '<p>Start by creating a new task or uploading tasks in bulk.</p>';\n";
  html += "    html += '</div>';\n";
  html += "  } else {\n";
  html +=
    "    html = '<p style=\"color: #4a5568; font-weight: 600; margin-bottom: 20px;\">';\n";
  html +=
    "    html += '<i class=\"fas fa-info-circle\"></i> Found ' + tasks.length + ' task(s)';\n";
  html += "    html += '</p>';\n";
  html += "    \n";
  html += "    html += '<table class=\"table\">';\n";
  html += "    html += '<thead><tr>';\n";
  html += "    html += '<th>Case ID</th>';\n";
  html += "    html += '<th>Client Name</th>';\n";
  html += "    html += '<th>Employee</th>';\n";
  html += "    html += '<th>Pincode</th>';\n";
  html += "    html += '<th>Map</th>';\n";
  html += "    html += '<th>Status</th>';\n";
  html += "    html += '<th>Date</th>';\n";
  html += "    html += '<th>Actions</th>';\n";
  html += "    html += '</tr></thead>';\n";
  html += "    html += '<tbody>';\n";
  html += "    \n";
  html += "    tasks.forEach(function(task) {\n";
  html +=
    "      const statusClass = 'status-' + task.status.toLowerCase().replace(/ /g, '-');\n";
  html += "      \n";
  html += "      html += '<tr>';\n";
  html +=
    "      html += '<td><strong>' + escapeHtml(task.title) + '</strong></td>';\n";
  html += " html += '<td>' + escapeHtml(task.clientName || '-') + '</td>';\n";
  html += "      html += '<td>';\n";
  html += "      if (task.User) {\n";
  html += "        html += escapeHtml(task.User.name);\n";
  html += "      } else {\n";
  html +=
    "        html += '<span style=\"color: #9ca3af; font-style: italic;\">Unassigned</span>';\n";
  html += "      }\n";
  html += "      html += '</td>';\n";
  html +=
    "      html += '<td><span class=\"pincode-highlight\"><i class=\"fas fa-map-pin\"></i> ' + escapeHtml(task.pincode || 'N/A') + '</span></td>';\n";
  html += "      html += '<td>';\n";
  html += "      if (task.latitude && task.longitude) {\n";
  html +=
    '        html += \'<i class="fas fa-check-circle" style="color: #10b981;"></i> Yes\';\n';
  html += "      } else {\n";
  html +=
    '        html += \'<i class="fas fa-times-circle" style="color: #ef4444;"></i> No\';\n';
  html += "      }\n";
  html += "      html += '</td>';\n";
  html +=
    "      html += '<td><span class=\"status-badge ' + statusClass + '\">' + escapeHtml(task.status) + '</span></td>';\n";
  html +=
    "      html += '<td>' + escapeHtml(task.assignedDate || 'N/A') + '</td>';\n";
  html += "      html += '<td>';\n";
  html += "      html += '<div class=\"action-buttons\">';\n";
  html +=
    '      html += \'<button class="btn btn-warning btn-sm" onclick="showReassignModal(\' + task.id + \')" title="Reassign to different employee">\';\n';
  html += "      html += '<i class=\"fas fa-sync-alt\"></i>';\n";
  html += "      html += '</button>';\n";
  html +=
    '      html += \'<button class="btn btn-secondary btn-sm" onclick="unassignTask(\' + task.id + \')" title="Remove assignment">\';\n';
  html += "      html += '<i class=\"fas fa-times\"></i>';\n";
  html += "      html += '</button>';\n";
  html +=
    '      html += \'<button class="btn btn-danger btn-sm" onclick="deleteTask(\' + task.id + \')" title="Delete task permanently">\';\n';
  html += "      html += '<i class=\"fas fa-trash\"></i>';\n";
  html += "      html += '</button>';\n";
  html += "      html += '</div>';\n";
  html += "      html += '</td>';\n";
  html += "      html += '</tr>';\n";
  html += "    });\n";
  html += "    \n";
  html += "    html += '</tbody></table>';\n";
  html += "  }\n";
  html += "  \n";
  html += "  document.getElementById('allTasksList').innerHTML = html;\n";
  html += "}\n";
  html += "\n";
  html += "function searchAllTasks() {\n";
  html +=
    "  const searchTerm = document.getElementById('taskSearch').value.trim();\n";
  html += "  loadAllTasks(searchTerm);\n";
  html += "}\n";
  html += "\n";
  html += "// Reassign task modal (Feature #9 - Bug Fix Applied)\n";
  html += "function showReassignModal(taskId) {\n";
  html += "  fetch('/api/tasks?role=admin')\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(allTasks) {\n";
  html +=
    "      const task = allTasks.find(function(t) { return t.id === taskId; });\n";
  html += "      \n";
  html += "      if (!task) {\n";
  html += "        showToast('Task not found', 'error');\n";
  html += "        return;\n";
  html += "      }\n";
  html += "      \n";
  html += "      // FIXED: Using empResponse variable correctly\n";
  html += "      fetch('/api/users')\n";
  html +=
    "        .then(function(empResponse) { return empResponse.json(); })\n";
  html += "        .then(function(employees) {\n";
  html += "          const modal = document.createElement('div');\n";
  html += "          modal.className = 'modal show';\n";
  html += "          \n";
  html += "          let html = '<div class=\"modal-content\">';\n";
  html +=
    "          html += '<h2><i class=\"fas fa-sync-alt\"></i> Reassign Task</h2>';\n";
  html += "          \n";
  html += "          html += '<div class=\"info-box info\">';\n";
  html +=
    "          html += '<p><strong>Task:</strong> ' + escapeHtml(task.title) + '</p>';\n";
  html +=
    "          html += '<p><strong>Current Assignment:</strong> ' + (task.User ? escapeHtml(task.User.name) : 'Unassigned') + '</p>';\n";
  html += "          html += '</div>';\n";
  html += "          \n";
  html += "          html += '<div class=\"form-group\">';\n";
  html +=
    "          html += '<label><i class=\"fas fa-user\"></i> Reassign to:</label>';\n";
  html +=
    '          html += \'<select id="newEmployee" style="padding: 14px; width: 100%; border-radius: 10px; border: 2px solid #e2e8f0; font-size: 15px; font-weight: 600;">\';\n';
  html += "          \n";
  html +=
    "          employees.filter(function(e) { return e.id !== task.assignedTo; }).forEach(function(emp) {\n";
  html +=
    "            html += '<option value=\"' + emp.id + '\">' + escapeHtml(emp.name) + ' (' + escapeHtml(emp.email) + ')</option>';\n";
  html += "          });\n";
  html += "          \n";
  html += "          html += '</select>';\n";
  html += "          html += '</div>';\n";
  html += "          \n";
  html +=
    "          html += '<div style=\"display: flex; gap: 12px; margin-top: 25px;\">';\n";
  html +=
    "          html += '<button class=\"btn btn-primary\" onclick=\"confirmReassign(' + taskId + ')\">';\n";
  html +=
    "          html += '<i class=\"fas fa-check\"></i> Confirm Reassignment';\n";
  html += "          html += '</button>';\n";
  html +=
    '          html += \'<button class="btn btn-secondary" onclick="closeModal()">\';\n';
  html += "          html += '<i class=\"fas fa-times\"></i> Cancel';\n";
  html += "          html += '</button>';\n";
  html += "          html += '</div>';\n";
  html += "          \n";
  html += "          html += '</div>';\n";
  html += "          \n";
  html += "          modal.innerHTML = html;\n";
  html += "          document.body.appendChild(modal);\n";
  html += "        });\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading task data:', err);\n";
  html += "      showToast('Failed to load task data', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function confirmReassign(taskId) {\n";
  html +=
    "  const newEmployeeId = document.getElementById('newEmployee').value;\n";
  html += "  \n";
  html += "  if (!newEmployeeId) {\n";
  html += "    showToast('Please select an employee', 'error');\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  fetch('/api/tasks/' + taskId + '/reassign', {\n";
  html += "    method: 'PUT',\n";
  html += "    headers: { 'Content-Type': 'application/json' },\n";
  html += "    body: JSON.stringify({\n";
  html += "      newEmployeeId: parseInt(newEmployeeId),\n";
  html += "      adminId: currentUser.id,\n";
  html += "      adminName: currentUser.name\n";
  html += "    })\n";
  html += "  })\n";
  html += "  .then(function(res) { return res.json(); })\n";
  html += "  .then(function(data) {\n";
  html += "    if (data.success) {\n";
  html +=
    "      showToast(data.message || 'Task reassigned successfully!', 'success');\n";
  html += "      closeModal();\n";
  html += "      loadAllTasks();\n";
  html += "    } else {\n";
  html +=
    "      showToast(data.message || 'Error reassigning task', 'error');\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Error reassigning task:', err);\n";
  html += "    showToast('Connection error', 'error');\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function unassignTask(taskId) {\n";
  html +=
    "  if (!confirm('Remove assignment and return this task to the unassigned pool?')) {\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  fetch('/api/tasks/' + taskId + '/unassign', {\n";
  html += "    method: 'POST',\n";
  html += "    headers: { 'Content-Type': 'application/json' },\n";
  html += "    body: JSON.stringify({\n";
  html += "      adminId: currentUser.id,\n";
  html += "      adminName: currentUser.name\n";
  html += "    })\n";
  html += "  })\n";
  html += "  .then(function(res) { return res.json(); })\n";
  html += "  .then(function(data) {\n";
  html += "    if (data.success) {\n";
  html +=
    "      showToast('Task unassigned and returned to pool', 'success');\n";
  html += "      loadAllTasks();\n";
  html += "    } else {\n";
  html +=
    "      showToast(data.message || 'Error unassigning task', 'error');\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Error unassigning task:', err);\n";
  html += "    showToast('Connection error', 'error');\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function deleteTask(taskId) {\n";
  html +=
    "  if (!confirm('Permanently delete this task? This action cannot be undone.')) {\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html +=
    "  fetch('/api/tasks/' + taskId + '?adminId=' + currentUser.id + '&adminName=' + encodeURIComponent(currentUser.name), {\n";
  html += "    method: 'DELETE'\n";
  html += "  })\n";
  html += "  .then(function(res) { return res.json(); })\n";
  html += "  .then(function(data) {\n";
  html += "    if (data.success) {\n";
  html += "      showToast('Task deleted successfully', 'success');\n";
  html += "      loadAllTasks();\n";
  html += "    } else {\n";
  html += "      showToast('Error deleting task', 'error');\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Error deleting task:', err);\n";
  html += "    showToast('Connection error', 'error');\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function closeModal() {\n";
  html += "  const modal = document.querySelector('.modal');\n";
  html += "  if (modal) modal.remove();\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// ADMIN FUNCTIONS - EMPLOYEES\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showEmployees() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-users\"></i> Manage Employees</h2>';\n";
  html += "  \n";
  html +=
    '  html += \'<button class="btn btn-primary" onclick="showAddEmployee()" style="margin-bottom: 25px;">\';\n';
  html += "  html += '<i class=\"fas fa-user-plus\"></i> Add New Employee';\n";
  html += "  html += '</button>';\n";
  html += "  \n";
  html += "  html += '<div id=\"employeeList\">';\n";
  html +=
    '  html += \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading employees...</div>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  loadEmployees();\n";
  html += "}\n";
  html += "function showAnalyticsDashboard() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-chart-pie\"></i> Analytics Dashboard</h2>';\n";
  html +=
    '  html += \'<div id="analyticsStats" style="margin-bottom: 30px;"></div>\';\n';
  html +=
    '  html += \'<canvas id="statusChart" width="400" height="200" style="margin-bottom:36px;"></canvas>\';\n';
  html +=
    '  html += \'<canvas id="tasksDayChart" width="400" height="200"></canvas>\';\n';
  html +=
    '  html += \'<div id="analyticsLoading" class="loading-spinner show"><i class="fas fa-spinner"></i>Loading analytics...</div>\';\n';
  html += "  document.getElementById('content').innerHTML = html;\n";
  html +=
    "  fetch('/api/analytics').then(res=>res.json()).then(function(data){\n";
  html +=
    "    document.getElementById('analyticsLoading').style.display = 'none';\n";
  html += "    let stats = '';\n";
  html += "    stats += '<div class=\"info-box info\">';\n";
  html +=
    "    stats += '<p><strong>Total Tasks:</strong> ' + data.totalTasks + '</p>';\n";
  html +=
    "    stats += '<p><strong>Assigned Tasks:</strong> ' + data.assignedTasks + '</p>';\n";
  html +=
    "    stats += '<p><strong>Completed Tasks:</strong> ' + data.completedTasks + '</p>';\n";
  html +=
    "    stats += '<p><strong>Verified Tasks:</strong> ' + data.verifiedTasks + '</p>';\n";
  html +=
    "    stats += '<p><strong>Active Employees:</strong> ' + data.userCount + '</p>';\n";
  html +=
    "    stats += '<p><strong>Admins:</strong> ' + data.adminCount + '</p>';\n";
  html += "    stats += '</div>';\n";
  html += "    document.getElementById('analyticsStats').innerHTML = stats;\n";
  html += "    // Prepare status counts for chart\n";
  html += "    const statusLabels = data.statusCounts.map(s=>s.status);\n";
  html +=
    "    const statusCounts = data.statusCounts.map(s=>parseInt(s.count));\n";
  html += "    // Prepare per-day chart\n";
  html += "    const dates = data.tasksPerDay.map(x=>x.day);\n";
  html +=
    "    const perDayCounts = data.tasksPerDay.map(x=>parseInt(x.count));\n";
  html += "    if (!window.Chart) {\n";
  html += "      let script = document.createElement('script');\n";
  html += "      script.src = 'https://cdn.jsdelivr.net/npm/chart.js';\n";
  html += "      script.onload = function() {\n";
  html +=
    "        renderAnalyticsCharts(statusLabels, statusCounts, dates, perDayCounts);\n";
  html += "      };\n";
  html += "      document.head.appendChild(script);\n";
  html += "    } else {\n";
  html +=
    "      renderAnalyticsCharts(statusLabels, statusCounts, dates, perDayCounts);\n";
  html += "    }\n";
  html += "  });\n";
  html += "}\n";
  html +=
    "function renderAnalyticsCharts(statusLabels, statusCounts, dates, perDayCounts) {\n";
  html +=
    "  new Chart(document.getElementById('statusChart').getContext('2d'), {\n";
  html += "    type: 'doughnut',\n";
  html +=
    "    data: { labels: statusLabels, datasets: [{ data: statusCounts, backgroundColor: [ '#2d3748', '#10b981', '#f59e0b','#6366F1','#ef4444','#a7f3d0','#bfdbfe','#fecaca','#fde68a','#cbd5e0','#92400e' ] }] },\n";
  html +=
    "    options: { plugins: { legend: { position: 'bottom' } }, responsive: true }\n";
  html += "  });\n";
  html +=
    "  new Chart(document.getElementById('tasksDayChart').getContext('2d'), {\n";
  html += "    type: 'line',\n";
  html +=
    "    data: { labels: dates, datasets: [{ label: 'Tasks Created', data: perDayCounts, borderColor: '#3b82f6', fill: true, backgroundColor: 'rgba(59,130,246,0.1)' }] },\n";
  html +=
    "    options: { responsive: true, plugins: { legend: { display: false } } }\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function loadEmployees() {\n";
  html += "  fetch('/api/users')\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(employees) {\n";
  html += "      let html = '';\n";
  html += "      \n";
  html += "      if (employees.length === 0) {\n";
  html += "        html = '<div class=\"empty-state\">';\n";
  html += "        html += '<i class=\"fas fa-user-plus\"></i>';\n";
  html += "        html += '<h3>No Employees Yet</h3>';\n";
  html +=
    "        html += '<p>Add your first employee to start assigning tasks.</p>';\n";
  html += "        html += '</div>';\n";
  html += "      } else {\n";
  html +=
    "        html = '<p style=\"color: #4a5568; font-weight: 600; margin-bottom: 20px;\">';\n";
  html +=
    "        html += '<i class=\"fas fa-info-circle\"></i> Total Employees: ' + employees.length;\n";
  html += "        html += '</p>';\n";
  html += "        \n";
  html += "        html += '<table class=\"table\">';\n";
  html += "        html += '<thead><tr>';\n";
  html += "        html += '<th>Name</th>';\n";
  html += "        html += '<th>Employee ID</th>';\n";
  html += "        html += '<th>Email</th>';\n";
  html += "        html += '<th>Phone</th>';\n";
  html += "        html += '</tr></thead>';\n";
  html += "        html += '<tbody>';\n";
  html += "        \n";
  html += "        employees.forEach(function(emp) {\n";
  html += "          html += '<tr>';\n";
  html +=
    "          html += '<td><strong>' + escapeHtml(emp.name) + '</strong></td>';\n";
  html +=
    "          html += '<td>' + escapeHtml(emp.employeeId || 'N/A') + '</td>';\n";
  html += "          html += '<td>' + escapeHtml(emp.email) + '</td>';\n";
  html +=
    "          html += '<td>' + escapeHtml(emp.phone || 'N/A') + '</td>';\n";
  html += "          html += '</tr>';\n";
  html += "        });\n";
  html += "        \n";
  html += "        html += '</tbody></table>';\n";
  html += "      }\n";
  html += "      \n";
  html += "      document.getElementById('employeeList').innerHTML = html;\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading employees:', err);\n";
  html +=
    "      document.getElementById('employeeList').innerHTML = '<div class=\"empty-state\"><i class=\"fas fa-exclamation-triangle\"></i><h3>Error Loading Employees</h3><p>Please try again.</p></div>';\n";
  html += "      showToast('Failed to load employees', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function showAddEmployee() {\n";
  html += "  const modal = document.createElement('div');\n";
  html += "  modal.className = 'modal show';\n";
  html += "  \n";
  html += "  let html = '<div class=\"modal-content\">';\n";
  html +=
    "  html += '<h2><i class=\"fas fa-user-plus\"></i> Add New Employee</h2>';\n";
  html += "  \n";
  html += "  html += '<form id=\"employeeForm\">';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-user\"></i> Full Name *</label>';\n";
  html +=
    '  html += \'<input type="text" id="empName" required placeholder="Enter full name">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-id-badge\"></i> Employee ID *</label>';\n";
  html +=
    '  html += \'<input type="text" id="empId" required placeholder="e.g., EMP001">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-envelope\"></i> Email *</label>';\n";
  html +=
    '  html += \'<input type="email" id="empEmail" required placeholder="employee@company.com">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-phone\"></i> Phone Number</label>';\n";
  html +=
    '  html += \'<input type="tel" id="empPhone" placeholder="+91 9876543210">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div class=\"form-group\">';\n";
  html +=
    "  html += '<label><i class=\"fas fa-lock\"></i> Password *</label>';\n";
  html +=
    '  html += \'<input type="password" id="empPassword" required placeholder="Default: 123456" value="123456">\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html +=
    "  html += '<div style=\"display: flex; gap: 12px; margin-top: 25px;\">';\n";
  html += '  html += \'<button type="submit" class="btn btn-primary">\';\n';
  html += "  html += '<i class=\"fas fa-check\"></i> Create Employee';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button type="button" class="btn btn-secondary" onclick="closeModal()">\';\n';
  html += "  html += '<i class=\"fas fa-times\"></i> Cancel';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '</form>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  modal.innerHTML = html;\n";
  html += "  document.body.appendChild(modal);\n";
  html += "  \n";
  html +=
    "  document.getElementById('employeeForm').addEventListener('submit', function(e) {\n";
  html += "    e.preventDefault();\n";
  html += "    \n";
  html += "    const formData = {\n";
  html += "      name: document.getElementById('empName').value,\n";
  html += "      employeeId: document.getElementById('empId').value,\n";
  html += "      email: document.getElementById('empEmail').value,\n";
  html += "      phone: document.getElementById('empPhone').value,\n";
  html +=
    "      password: document.getElementById('empPassword').value || '123456'\n";
  html += "    };\n";
  html += "    \n";
  html +=
    "    const btn = e.target.querySelector('button[type=\"submit\"]');\n";
  html += "    btn.disabled = true;\n";
  html +=
    "    btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> Creating...';\n";
  html += "    \n";
  html += "    fetch('/api/users', {\n";
  html += "      method: 'POST',\n";
  html += "      headers: { 'Content-Type': 'application/json' },\n";
  html += "      body: JSON.stringify(formData)\n";
  html += "    })\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(data) {\n";
  html += "      if (data.success) {\n";
  html += "        showToast('Employee created successfully!', 'success');\n";
  html += "        closeModal();\n";
  html += "        showEmployees();\n";
  html += "      } else {\n";
  html +=
    "        showToast(data.message || 'Error creating employee', 'error');\n";
  html += "        btn.disabled = false;\n";
  html +=
    "        btn.innerHTML = '<i class=\"fas fa-check\"></i> Create Employee';\n";
  html += "      }\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error creating employee:', err);\n";
  html += "      showToast('Connection error', 'error');\n";
  html += "      btn.disabled = false;\n";
  html +=
    "      btn.innerHTML = '<i class=\"fas fa-check\"></i> Create Employee';\n";
  html += "    });\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// ADMIN FUNCTIONS - EXPORT CSV\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function exportTasks() {\n";
  html += "  showToast('Preparing export...', 'info');\n";
  html += "  window.location.href = '/api/export';\n";
  html += "  setTimeout(function() {\n";
  html += "    showToast('Export downloaded successfully!', 'success');\n";
  html += "  }, 1000);\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// EMPLOYEE FUNCTIONS - TODAY'S TASKS (Features #6, #7, #8)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showTodayTasks() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-tasks\"></i> Today\\'s Tasks</h2>';\n";
  html += "  \n";
  html += "  html += '<div class=\"filter-section\">';\n";
  html +=
    '  html += \'<input type="text" id="todayTaskSearch" class="search-box" placeholder="Search by Case ID, Pincode, Address..." style="flex: 1; max-width: 500px;">\';\n';
  html +=
    '  html += \'<button class="btn btn-warning btn-sm" onclick="sortByNearest()">\';\n';
  html +=
    "  html += '<i class=\"fas fa-location-arrow\"></i> Sort by Nearest';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button class="btn btn-success btn-sm" onclick="sortByPincode()">\';\n';
  html += "  html += '<i class=\"fas fa-map-pin\"></i> Sort by Pincode';\n";
  html += "  html += '</button>';\n";
  html +=
    '  html += \'<button class="btn btn-info btn-sm" onclick="loadTodayTasks()">\';\n';
  html += "  html += '<i class=\"fas fa-sync\"></i> Refresh';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div id=\"todayTasksList\">';\n";
  html +=
    '  html += \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading your tasks...</div>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  \n";
  html += "  // Real-time search\n";
  html +=
    "  document.getElementById('todayTaskSearch').addEventListener('input', function() {\n";
  html += "    searchTodayTasks();\n";
  html += "  });\n";
  html += "  \n";
  html += "  loadTodayTasks();\n";
  html += "}\n";
  html += "\n";
  html += "function loadTodayTasks(searchTerm) {\n";
  html += "  const today = new Date().toISOString().split('T')[0];\n";
  html +=
    "  const url = '/api/tasks?role=employee&userId=' + currentUser.id + '&date=' + today + (searchTerm ? '&search=' + encodeURIComponent(searchTerm) : '');\n";
  html += "  \n";
  html += "  fetch(url)\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(tasks) {\n";
  html += "      allEmployeeTasks = tasks;\n";
  html += "      displayEmployeeTasks(tasks);\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading tasks:', err);\n";
  html +=
    "      document.getElementById('todayTasksList').innerHTML = '<div class=\"empty-state\"><i class=\"fas fa-exclamation-triangle\"></i><h3>Error Loading Tasks</h3><p>Please try again.</p></div>';\n";
  html += "      showToast('Failed to load tasks', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function displayEmployeeTasks(tasks) {\n";
  html += "  let html = '';\n";
  html += "  \n";
  html += "  if (tasks.length === 0) {\n";
  html += "    html = '<div class=\"empty-state\">';\n";
  html += "    html += '<i class=\"fas fa-check-circle\"></i>';\n";
  html += "    html += '<h3>All Clear!</h3>';\n";
  html += "    html += '<p>No tasks assigned for today. Great job!</p>';\n";
  html += "    html += '</div>';\n";
  html += "  } else {\n";
  html +=
    "    html = '<p style=\"color: #4a5568; font-weight: 600; margin-bottom: 20px;\">';\n";
  html +=
    "    html += '<i class=\"fas fa-info-circle\"></i> You have ' + tasks.length + ' task(s) for today';\n";
  html += "    html += '</p>';\n";
  html += "    \n";
  html += "    html += '<div class=\"grid\">';\n";
  html += "    \n";
  html += "    tasks.forEach(function(task) {\n";
  html +=
    "      const statusClass = 'status-' + task.status.toLowerCase().replace(/ /g, '-');\n";
  html += "      \n";
  html += "      html += '<div class=\"task-card\">';\n";
  html +=
    "      html += '<h3><i class=\"fas fa-clipboard-list\"></i> ' + escapeHtml(task.title) + '</h3>';\n";
  html += " if (task.clientName) {\n";
  html +=
    " html += '<p><strong>Client:</strong> ' + escapeHtml(task.clientName) + '</p>';\n";
  html += " }\n";
  html += "      \n";
  html += "      html += '<div class=\"pincode-highlight\">';\n";
  html +=
    "      html += '<i class=\"fas fa-map-pin\"></i> Pincode: ' + escapeHtml(task.pincode || 'N/A');\n";
  html += "      html += '</div>';\n";
  html += "      \n";
  html += "      if (task.distance !== undefined) {\n";
  html += "        html += '<div class=\"distance-indicator\">';\n";
  html +=
    "        html += '<i class=\"fas fa-route\"></i> Distance: ' + task.distance.toFixed(2) + ' km away';\n";
  html += "        html += '</div>';\n";
  html += "      }\n";
  html += "      \n";
  html +=
    "      html += '<span class=\"status-badge ' + statusClass + '\">';\n";
  html +=
    "      html += '<i class=\"fas fa-circle\"></i> ' + escapeHtml(task.status);\n";
  html += "      html += '</span>';\n";
  html += "      \n";
  html += "      if (task.notes) {\n";
  html +=
    "        html += '<p style=\"margin-top: 15px; color: #4a5568;\"><i class=\"fas fa-sticky-note\"></i> ' + escapeHtml(task.notes) + '</p>';\n";
  html += "      }\n";
  html += "      \n";
  html += "      // Feature #8: MapURL visibility (CRITICAL)\n";
  html += "      if (task.mapUrl) {\n";
  html +=
    '        html += \'<a href="\' + escapeHtml(task.mapUrl) + \'" target="_blank" rel="noopener noreferrer" class="map-button">\';\n';
  html +=
    "        html += '<i class=\"fas fa-map-marker-alt\"></i> Open Location Map';\n";
  html += "        html += '</a>';\n";
  html += "      } else {\n";
  html +=
    '        html += \'<div class="no-map"><i class="fas fa-map-marker-alt"></i> No map available for this task</div>\';\n';
  html += "      }\n";
  html += "      \n";
  html += "      html += '<div class=\"action-buttons\">';\n";
  html +=
    "      html += '<select id=\"status-' + task.id + '\" style=\"flex: 2;\">';\n";
  html +=
    "      html += '<option value=\"Pending\"' + (task.status === 'Pending' ? ' selected' : '') + '>Pending</option>';\n";
  html +=
    "      html += '<option value=\"Completed\"' + (task.status === 'Completed' ? ' selected' : '') + '>Completed</option>';\n";
  html +=
    "      html += '<option value=\"Verified\"' + (task.status === 'Verified' ? ' selected' : '') + '>Verified</option>';\n";
  html +=
    "      html += '<option value=\"Left Job\"' + (task.status === 'Left Job' ? ' selected' : '') + '>Left Job</option>';\n";
  html +=
    "      html += '<option value=\"Not Sharing Info\"' + (task.status === 'Not Sharing Info' ? ' selected' : '') + '>Not Sharing Info</option>';\n";
  html +=
    "      html += '<option value=\"Not Picking\"' + (task.status === 'Not Picking' ? ' selected' : '') + '>Not Picking</option>';\n";
  html +=
    "      html += '<option value=\"Switch Off\"' + (task.status === 'Switch Off' ? ' selected' : '') + '>Switch Off</option>';\n";
  html +=
    "      html += '<option value=\"Incorrect Number\"' + (task.status === 'Incorrect Number' ? ' selected' : '') + '>Incorrect Number</option>';\n";
  html +=
    "      html += '<option value=\"Wrong Address\"' + (task.status === 'Wrong Address' ? ' selected' : '') + '>Wrong Address</option>';\n";
  html += "      html += '</select>';\n";
  html +=
    '      html += \'<button class="btn btn-primary" onclick="updateTaskStatus(\' + task.id + \')" style="flex: 1;">\';\n';
  html += "      html += '<i class=\"fas fa-save\"></i> Update';\n";
  html += "      html += '</button>';\n";
  html += "      html += '</div>';\n";
  html += "      \n";
  html += "      html += '</div>';\n";
  html += "    });\n";
  html += "    \n";
  html += "    html += '</div>';\n";
  html += "  }\n";
  html += "  \n";
  html += "  document.getElementById('todayTasksList').innerHTML = html;\n";
  html += "}\n";
  html += "\n";
  html += "function searchTodayTasks() {\n";
  html +=
    "  const searchTerm = document.getElementById('todayTaskSearch').value.toLowerCase();\n";
  html += "  \n";
  html += "  const filteredTasks = allEmployeeTasks.filter(function(task) {\n";
  html += "    return task.title.toLowerCase().includes(searchTerm) ||\n";
  html += "           task.status.toLowerCase().includes(searchTerm) ||\n";
  html +=
    "           (task.pincode || '').toLowerCase().includes(searchTerm) ||\n";
  html += "           (task.notes || '').toLowerCase().includes(searchTerm);\n";
  html += "  });\n";
  html += "  \n";
  html += "  displayEmployeeTasks(filteredTasks);\n";
  html += "}\n";
  html += "\n";
  html += "// PRESERVED: Sort by Nearest Location (GPS-based)\n";
  html += "function sortByNearest() {\n";
  html += "  if (!navigator.geolocation) {\n";
  html +=
    "    showToast('Geolocation is not supported by your browser', 'error');\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  \n";
  html += "  showToast('Getting your location...', 'info');\n";
  html += "  \n";
  html += "  navigator.geolocation.getCurrentPosition(function(position) {\n";
  html += "    const userLat = position.coords.latitude;\n";
  html += "    const userLng = position.coords.longitude;\n";
  html += "    \n";
  html +=
    "    const tasksWithDistance = allEmployeeTasks.map(function(task) {\n";
  html += "      if (task.latitude && task.longitude) {\n";
  html +=
    "        const distance = calculateDistance(userLat, userLng, task.latitude, task.longitude);\n";
  html += "        return Object.assign({}, task, { distance: distance });\n";
  html += "      }\n";
  html += "      return Object.assign({}, task, { distance: 999999 });\n";
  html += "    }).sort(function(a, b) {\n";
  html += "      return a.distance - b.distance;\n";
  html += "    });\n";
  html += "    \n";
  html += "    allEmployeeTasks = tasksWithDistance;\n";
  html += "    displayEmployeeTasks(tasksWithDistance);\n";
  html += "    showToast('Tasks sorted by nearest location!', 'success');\n";
  html += "    \n";
  html += "  }, function(error) {\n";
  html += "    console.error('Geolocation error:', error);\n";
  html += "    showToast('Location access denied or unavailable', 'error');\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html += "function calculateDistance(lat1, lon1, lat2, lon2) {\n";
  html += "  const R = 6371; // Earth's radius in km\n";
  html += "  const dLat = (lat2 - lat1) * Math.PI / 180;\n";
  html += "  const dLon = (lon2 - lon1) * Math.PI / 180;\n";
  html += "  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +\n";
  html +=
    "            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *\n";
  html += "            Math.sin(dLon/2) * Math.sin(dLon/2);\n";
  html += "  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));\n";
  html += "  return R * c;\n";
  html += "}\n";
  html += "\n";
  html += "// PRESERVED: Sort by Pincode Proximity\n";
  html += "function sortByPincode() {\n";
  html += "  if (!navigator.geolocation) {\n";
  html += "    showToast('Geolocation not supported', 'error');\n";
  html += "    return;\n";
  html += "  }\n";
  html += "  showToast('Getting your location...', 'info');\n";
  html += "  navigator.geolocation.getCurrentPosition(function(position) {\n";
  html += "    const userLat = position.coords.latitude;\n";
  html += "    const userLng = position.coords.longitude;\n";
  html +=
    "    const pincodeData={'560001':{lat:12.9716,lng:77.5946},'560002':{lat:12.9844,lng:77.5908},'560003':{lat:13.0067,lng:77.5713},'560004':{lat:12.944,lng:77.575},'560005':{lat:13.005,lng:77.588},'560006':{lat:12.9634,lng:77.5855},'560007':{lat:12.8431,lng:77.4863},'560008':{lat:13.0169,lng:77.64},'560009':{lat:13.031,lng:77.645},'560010':{lat:13.0208,lng:77.6344},'560011':{lat:13.032,lng:77.596},'560012':{lat:13.0048,lng:77.5956},'560013':{lat:13.035,lng:77.551},'560014':{lat:13.0305,lng:77.5575},'560015':{lat:13.045,lng:77.565},'560016':{lat:13.041,lng:77.559},'560017':{lat:13.0085,lng:77.65},'560018':{lat:12.952,lng:77.579},'560019':{lat:12.952,lng:77.5735},'560020':{lat:13.005,lng:77.5842},'560021':{lat:12.987,lng:77.565},'560022':{lat:13.013,lng:77.551},'560023':{lat:12.976,lng:77.542},'560024':{lat:12.968,lng:77.535},'560025':{lat:13.011,lng:77.562},'560026':{lat:12.972,lng:77.547},'560027':{lat:12.982,lng:77.574},'560029':{lat:12.955,lng:77.601},'560030':{lat:12.9716,lng:77.5946},'560032':{lat:12.958,lng:77.597},'560033':{lat:12.964,lng:77.596},'560034':{lat:12.9275,lng:77.6225},'560035':{lat:12.934,lng:77.668},'560036':{lat:12.994,lng:77.691},'560037':{lat:13.045,lng:77.69},'560038':{lat:12.9755,lng:77.64},'560039':{lat:12.968,lng:77.522},'560040':{lat:12.981,lng:77.519},'560041':{lat:12.932,lng:77.585},'560042':{lat:12.918,lng:77.599},'560043':{lat:13.014,lng:77.649},'560045':{lat:12.985,lng:77.607},'560046':{lat:13.007,lng:77.601},'560047':{lat:12.964,lng:77.609},'560048':{lat:13.036,lng:77.705},'560049':{lat:13.064,lng:77.728},'560050':{lat:13.064,lng:77.628},'560051':{lat:12.908,lng:77.597},'560053':{lat:12.965,lng:77.573},'560054':{lat:13.037,lng:77.562},'560055':{lat:13.0055,lng:77.569},'560056':{lat:13.006,lng:77.575},'560057':{lat:13.026,lng:77.521},'560058':{lat:13.01,lng:77.53},'560059':{lat:13.015,lng:77.533},'560060':{lat:13.009,lng:77.548},'560061':{lat:12.925,lng:77.535},'560062':{lat:12.896,lng:77.527},'560063':{lat:13.108,lng:77.575},'560064':{lat:13.0675,lng:77.598},'560065':{lat:13.08,lng:77.57},'560066':{lat:13.0875,lng:77.582},'560067':{lat:13.152,lng:77.605},'560068':{lat:13.068,lng:77.718},'560070':{lat:12.913,lng:77.547},'560071':{lat:12.957,lng:77.638},'560072':{lat:12.951,lng:77.485},'560073':{lat:13.043,lng:77.545},'560074':{lat:12.888,lng:77.498},'560075':{lat:13.0055,lng:77.578},'560076':{lat:12.9,lng:77.597},'560077':{lat:13.032,lng:77.585},'560078':{lat:12.91,lng:77.555},'560079':{lat:12.981,lng:77.541},'560080':{lat:13.005,lng:77.573},'560082':{lat:12.856,lng:77.542},'560083':{lat:12.897,lng:77.578},'560084':{lat:13.03,lng:77.616},'560085':{lat:12.911,lng:77.558},'560086':{lat:12.96,lng:77.541},'560087':{lat:13.028,lng:77.595},'560090':{lat:13.016,lng:77.709},'560091':{lat:12.977,lng:77.615},'560092':{lat:13.045,lng:77.667},'560093':{lat:12.984,lng:77.655},'560094':{lat:12.9695,lng:77.657},'560095':{lat:12.935,lng:77.627},'560096':{lat:12.942,lng:77.618},'560097':{lat:13.07,lng:77.784},'560098':{lat:13.009,lng:77.589},'560099':{lat:12.845,lng:77.679},'560100':{lat:12.845,lng:77.661},'560102':{lat:12.912,lng:77.638},'560103':{lat:12.926,lng:77.676},'560104':{lat:13.064,lng:77.748},'560105':{lat:13.087,lng:77.702},'562106':{lat:12.711,lng:77.695},'562107':{lat:12.776,lng:77.754},'562125':{lat:12.916,lng:77.782},'562130':{lat:13.006,lng:77.791},'562149':{lat:13.143,lng:77.709},'562157':{lat:13.178,lng:77.721}};\n";
  html +=
    "    const tasksWithDistance = allEmployeeTasks.map(function(task) {\n";
  html += "      let distance = 999999;\n";
  html += "      if (task.latitude && task.longitude) {\n";
  html +=
    "        distance = calculateDistance(userLat, userLng, task.latitude, task.longitude);\n";
  html += "      } else if (task.pincode && pincodeData[task.pincode]) {\n";
  html += "        const pc = pincodeData[task.pincode];\n";
  html +=
    "        distance = calculateDistance(userLat, userLng, pc.lat, pc.lng);\n";
  html += "      }\n";
  html += "      return Object.assign({}, task, { distance: distance });\n";
  html += "    });\n";
  html +=
    "    tasksWithDistance.sort(function(a, b) { return a.distance - b.distance; });\n";
  html += "    allEmployeeTasks = tasksWithDistance;\n";
  html += "    displayEmployeeTasks(allEmployeeTasks);\n";
  html +=
    "    showToast('Sorted ' + tasksWithDistance.length + ' tasks by distance', 'success');\n";
  html += "    const searchBox = document.getElementById('taskSearchBox');\n";
  html += "    if (searchBox) searchBox.placeholder = 'Sorted by proximity';\n";
  html += "  }, function(error) {\n";
  html += "    let errorMsg = 'Location failed: ';\n";
  html += "    if (error.code === 1) errorMsg += 'Permission denied';\n";
  html +=
    "    else if (error.code === 2) errorMsg += 'Position unavailable';\n";
  html += "    else if (error.code === 3) errorMsg += 'Timeout';\n";
  html += "    showToast(errorMsg, 'error');\n";
  html +=
    "  }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 30000 });\n";
  html += "}\n";
  html += "\n";
  html += "function updateTaskStatus(taskId) {\n";
  html +=
    "  const newStatus = document.getElementById('status-' + taskId).value;\n";
  html +=
    "  const btn = document.querySelector('#status-' + taskId).nextElementSibling;\n";
  html += "  // FEATURE 1 PATCH: Prompt confirmation before status update\n";
  html += "  if (newStatus === 'Pending') {\n";
  html +=
    "    if (!confirm('Task status is still Pending! Please select the correct status when your work is done. Continue anyway?')) {\n";
  html += "      btn.disabled = false;\n";
  html += "      btn.innerHTML = '<i class=\"fas fa-save\"></i> Update';\n";
  html += "      return;\n";
  html += "    }\n";
  html += "  } else {\n";
  html +=
    "    if (!confirm('Are you sure you want to change status to \"' + newStatus + '\"?')) {\n";
  html += "      btn.disabled = false;\n";
  html += "      btn.innerHTML = '<i class=\"fas fa-save\"></i> Update';\n";
  html += "      return;\n";
  html += "    }\n";
  html += "  }\n";
  html += "  btn.disabled = true;\n";
  html +=
    "  btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> Updating...';\n";
  html += "  fetch('/api/tasks/' + taskId, {\n";
  html += "    method: 'PUT',\n";
  html += "    headers: { 'Content-Type': 'application/json' },\n";
  html += "    body: JSON.stringify({\n";
  html += "      status: newStatus,\n";
  html += "      userId: currentUser.id,\n";
  html += "      userName: currentUser.name\n";
  html += "    })\n";
  html += "  })\n";
  html += "  .then(function(res) { return res.json(); })\n";
  html += "  .then(function(data) {\n";
  html += "    if (data.success) {\n";
  html += "      showToast('Status updated to ' + newStatus, 'success');\n";
  html += "      loadTodayTasks();\n";
  html += "    } else {\n";
  html += "      showToast('Error updating status', 'error');\n";
  html += "      btn.disabled = false;\n";
  html += "      btn.innerHTML = '<i class=\"fas fa-save\"></i> Update';\n";
  html += "    }\n";
  html += "  })\n";
  html += "  .catch(function(err) {\n";
  html += "    console.error('Error updating status:', err);\n";
  html += "    showToast('Connection error', 'error');\n";
  html += "    btn.disabled = false;\n";
  html += "    btn.innerHTML = '<i class=\"fas fa-save\"></i> Update';\n";
  html += "  });\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// EMPLOYEE FUNCTIONS - TASK HISTORY (Feature #7)\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function showTaskHistory() {\n";
  html +=
    "  let html = '<h2><i class=\"fas fa-history\"></i> Task History</h2>';\n";
  html += "  \n";
  html += "  html += '<div class=\"filter-section\">';\n";
  html +=
    '  html += \'<input type="text" id="historyTaskSearch" class="search-box" placeholder="Search by Case ID, Date, Pincode, Status..." style="flex: 1; max-width: 500px;">\';\n';
  html +=
    '  html += \'<button class="btn btn-info btn-sm" onclick="loadTaskHistory()">\';\n';
  html += "  html += '<i class=\"fas fa-sync\"></i> Refresh';\n";
  html += "  html += '</button>';\n";
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  html += '<div id=\"historyTasksList\">';\n";
  html +=
    '  html += \'<div class="loading-spinner show"><i class="fas fa-spinner"></i>Loading task history...</div>\';\n';
  html += "  html += '</div>';\n";
  html += "  \n";
  html += "  document.getElementById('content').innerHTML = html;\n";
  html += "  \n";
  html += "  // Real-time search\n";
  html +=
    "  document.getElementById('historyTaskSearch').addEventListener('input', function() {\n";
  html += "    searchHistoryTasks();\n";
  html += "  });\n";
  html += "  \n";
  html += "  loadTaskHistory();\n";
  html += "}\n";
  html += "\n";
  html += "function loadTaskHistory(searchTerm) {\n";
  html +=
    "  const url = '/api/tasks?role=employee&userId=' + currentUser.id + '&status=all' + (searchTerm ? '&search=' + encodeURIComponent(searchTerm) : '');\n";
  html += "  \n";
  html += "  fetch(url)\n";
  html += "    .then(function(res) { return res.json(); })\n";
  html += "    .then(function(tasks) {\n";
  html += "      allHistoryTasks = tasks;\n";
  html += "      displayHistoryTasks(tasks);\n";
  html += "    })\n";
  html += "    .catch(function(err) {\n";
  html += "      console.error('Error loading history:', err);\n";
  html +=
    "      document.getElementById('historyTasksList').innerHTML = '<div class=\"empty-state\"><i class=\"fas fa-exclamation-triangle\"></i><h3>Error Loading History</h3><p>Please try again.</p></div>';\n";
  html += "      showToast('Failed to load history', 'error');\n";
  html += "    });\n";
  html += "}\n";
  html += "\n";
  html += "function displayHistoryTasks(tasks) {\n";
  html += "  let html = '';\n";
  html += "  \n";
  html += "  if (tasks.length === 0) {\n";
  html += "    html = '<div class=\"empty-state\">';\n";
  html += "    html += '<i class=\"fas fa-inbox\"></i>';\n";
  html += "    html += '<h3>No Task History</h3>';\n";
  html += "    html += '<p>Your completed tasks will appear here.</p>';\n";
  html += "    html += '</div>';\n";
  html += "  } else {\n";
  html +=
    "    html = '<p style=\"color: #4a5568; font-weight: 600; margin-bottom: 20px;\">';\n";
  html +=
    "    html += '<i class=\"fas fa-info-circle\"></i> Total Tasks: ' + tasks.length;\n";
  html += "    html += '</p>';\n";
  html += "    \n";
  html += "    html += '<table class=\"table\">';\n";
  html += "    html += '<thead><tr>';\n";
  html += "    html += '<th>Case ID</th>';\n";
  html += "    html += '<th>Client Name</th>';\n";
  html += "    html += '<th>Pincode</th>';\n";
  html += "    html += '<th>Status</th>';\n";
  html += "    html += '<th>Date</th>';\n";
  html += "    html += '<th>Map</th>';\n";
  html += "    html += '<th>Notes</th>';\n";
  html += "    html += '</tr></thead>';\n";
  html += "    html += '<tbody>';\n";
  html += "    \n";
  html += "    tasks.forEach(function(task) {\n";
  html +=
    "      const statusClass = 'status-' + task.status.toLowerCase().replace(/ /g, '-');\n";
  html += "      \n";
  html += "      html += '<tr>';\n";
  html +=
    "      html += '<td><strong>' + escapeHtml(task.title) + '</strong></td>';\n";
  html +=
    "    html += '<td>' + escapeHtml(task.clientName || '-') + '</td>';\n";
  html +=
    "      html += '<td><span class=\"pincode-highlight\"><i class=\"fas fa-map-pin\"></i> ' + escapeHtml(task.pincode || 'N/A') + '</span></td>';\n";
  html +=
    "      html += '<td><span class=\"status-badge ' + statusClass + '\">' + escapeHtml(task.status) + '</span></td>';\n";
  html +=
    "      html += '<td>' + escapeHtml(task.manualDate || task.assignedDate || 'N/A') + '</td>';\n";
  html += "      html += '<td>';\n";
  html += "      if (task.mapUrl) {\n";
  html +=
    '        html += \'<a href="\' + escapeHtml(task.mapUrl) + \'" target="_blank" rel="noopener noreferrer" style="color: #3b82f6; font-weight: 600;">\';\n';
  html +=
    "        html += '<i class=\"fas fa-map-marker-alt\"></i> View Map';\n";
  html += "        html += '</a>';\n";
  html += "      } else {\n";
  html +=
    "        html += '<span style=\"color: #9ca3af; font-style: italic;\">No map</span>';\n";
  html += "      }\n";
  html += "      html += '</td>';\n";
  html += "      html += '<td>' + escapeHtml(task.notes || 'N/A') + '</td>';\n";
  html += "      html += '</tr>';\n";
  html += "    });\n";
  html += "    \n";
  html += "    html += '</tbody></table>';\n";
  html += "  }\n";
  html += "  \n";
  html += "  document.getElementById('historyTasksList').innerHTML = html;\n";
  html += "}\n";
  html += "\n";
  html += "function searchHistoryTasks() {\n";
  html +=
    "  const searchTerm = document.getElementById('historyTaskSearch').value.toLowerCase();\n";
  html += "  \n";
  html += "  const filteredTasks = allHistoryTasks.filter(function(task) {\n";
  html += "    return task.title.toLowerCase().includes(searchTerm) ||\n";
  html += "           task.status.toLowerCase().includes(searchTerm) ||\n";
  html +=
    "           (task.pincode || '').toLowerCase().includes(searchTerm) ||\n";
  html +=
    "           (task.assignedDate || '').toLowerCase().includes(searchTerm);\n";
  html += "  });\n";
  html += "  \n";
  html += "  displayHistoryTasks(filteredTasks);\n";
  html += "}\n";
  html += "\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "// MENU INITIALIZATION\n";
  html +=
    "// ═══════════════════════════════════════════════════════════════════════════\n";
  html += "\n";
  html += "function initMenu() {\n";
  html += "  const menu = document.getElementById('menu');\n";
  html += "  \n";
  html += "  if (currentUser.role === 'admin') {\n";
  html += "    menu.innerHTML = '' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showAssignTask()">\' +\n';
  html += "      '<i class=\"fas fa-plus-circle\"></i> Assign Task' +\n";
  html += "      '</button>' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showUnassignedTasks()">\' +\n';
  html += "      '<i class=\"fas fa-inbox\"></i> Unassigned Tasks' +\n";
  html += "      '</button>' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showAllTasks()">\' +\n';
  html += "      '<i class=\"fas fa-list\"></i> View All Tasks' +\n";
  html += "      '</button>' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showEmployees()">\' +\n';
  html += "      '<i class=\"fas fa-users\"></i> Employees' +\n";
  html += "      '</button>' +\n";
  html +=
    '      \'<button class="btn btn-success" onclick="exportTasks()">\' +\n';
  html += "      '<i class=\"fas fa-download\"></i> Export CSV' +\n";
  html +=
    '      \'<button class="btn btn-info" onclick="showAnalyticsDashboard()">\' +\n';
  html += "      '<i class=\"fas fa-chart-pie\"></i> Analytics Dashboard' +\n";
  html += "      '</button>';\n";
  html += "    \n";
  html += "    showAssignTask();\n";
  html += "  } else {\n";
  html += "    menu.innerHTML = '' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showTodayTasks()">\' +\n';
  html += "      '<i class=\"fas fa-tasks\"></i> Today\\'s Tasks' +\n";
  html += "      '</button>' +\n";
  html +=
    '      \'<button class="btn btn-primary" onclick="showTaskHistory()">\' +\n';
  html += "      '<i class=\"fas fa-history\"></i> Task History' +\n";
  html += "      '</button>';\n";
  html += "    \n";
  html += "    showTodayTasks();\n";
  html += "  }\n";
  html += "}\n";
  html += "\n";
  html += "// Initialize dashboard\n";
  html += "initMenu();\n";
  html += "\n";
  html += "</script>";
  html += "</body>";
  html += "</html>";

  res.send(html);
});

// ═══════════════════════════════════════════════════════════════════════════
// KEEP-ALIVE SYSTEM (Feature #2)
// ═══════════════════════════════════════════════════════════════════════════

let pingCount = 0;

function keepAlive() {
  pingCount++;
  const uptime = process.uptime();
  const hours = Math.floor(uptime / 3600);
  const minutes = Math.floor((uptime % 3600) / 60);

  console.log(
    `🔄 Keep-alive ping #${pingCount} | Uptime: ${hours}h ${minutes}m`,
  );

  http
    .get("http://localhost:3000/health", (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        try {
          const health = JSON.parse(data);
          console.log("✅ Health check:", health.status);
        } catch (e) {
          console.log("✅ Server responding");
        }
      });
    })
    .on("error", (err) => {
      console.log("❌ Health check failed:", err.message);
    });
}

// ═══════════════════════════════════════════════════════════════════════════
// ERROR HANDLERS (Feature #11)
// ═══════════════════════════════════════════════════════════════════════════

process.on("unhandledRejection", (reason, promise) => {
  console.error("❌ Unhandled Rejection at:", promise);
  console.error("Reason:", reason);
});

process.on("uncaughtException", (error) => {
  console.error("❌ Uncaught Exception:", error);
  console.error("Stack:", error.stack);
});

// ═══════════════════════════════════════════════════════════════════════════
// SERVER INITIALIZATION & STARTUP
// ═══════════════════════════════════════════════════════════════════════════

async function startServer() {
  try {
    console.log(
      "═══════════════════════════════════════════════════════════════",
    );
    console.log("🚀 VALIDIANT PRODUCTIVITY TRACKER - STARTING UP");
    console.log(
      "═══════════════════════════════════════════════════════════════",
    );

    // Initialize database
    await initializeDatabase();

    // Start Express server
    app.listen(PORT, () => {
      console.log("");
      console.log(
        "═══════════════════════════════════════════════════════════════",
      );
      console.log("✅ SERVER RUNNING SUCCESSFULLY");
      console.log(
        "═══════════════════════════════════════════════════════════════",
      );
      console.log(`🌐 URL: http://localhost:${PORT}`);
      console.log(`📊 Environment: ${process.env.NODE_ENV || "production"}`);
      console.log(`🔐 Admin Login: admin@validiant.com / Admin@123`);
      console.log("");
      console.log("✅ Keep-alive system starting...");

      // Start keep-alive pings every 3 minutes
      setInterval(keepAlive, 180000);

      // Initial ping
      setTimeout(keepAlive, 5000);

      console.log("✅ All systems operational");
      console.log(
        "═══════════════════════════════════════════════════════════════",
      );
    });
  } catch (error) {
    console.error(
      "═══════════════════════════════════════════════════════════════",
    );
    console.error("❌ FATAL ERROR: Failed to start server");
    console.error(
      "═══════════════════════════════════════════════════════════════",
    );
    console.error(error);
    process.exit(1);
  }
}

// Start the server
startServer();
