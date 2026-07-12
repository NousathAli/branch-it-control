import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { doc, getDoc, setDoc, serverTimestamp } from "firebase/firestore";
import { db, ensureFirebaseLogin } from "./firebase";

const defaultUsers = [
  { id: 1, username: "demo-admin", password: "DemoAdmin@123", role: "admin", displayName: "Demo Admin" },
  { id: 2, username: "demo-viewer", password: "DemoViewer@123", role: "user", displayName: "Demo Viewer" },
  { id: 3, username: "demo-tech", password: "DemoTech@123", role: "user", displayName: "Demo Technician" },
];

const defaultUserProfiles = [
  {
    name: "Demo Technician",
    username: "demo-tech",
    roleTitle: "IT Support & CCTV Technician",
    department: "IT Department",
    company: "Demo Restaurant Group",
    reportingTo: "Demo Team Lead - Team Leader / IT Manager",
    specialization: "CCTV, NVR, IT infrastructure, branch opening setup, POS/network support",
    locationScope: "Company-owned F&B branches across UAE",
    profileSummary: "Handles branch IT setup, CCTV/NVR installation and review, network infrastructure, POS setup, IT rack installation, asset documentation, handover, pullout and fault follow-up.",
  },
  {
    name: "Demo Engineer",
    username: "demo-engineer",
    roleTitle: "IT Support / CCTV & Infrastructure Support",
    department: "IT Department",
    company: "Demo Restaurant Group",
    reportingTo: "Demo Team Lead - Team Leader / IT Manager",
    specialization: "CCTV support, branch infrastructure work, site support and IT asset handling",
    locationScope: "Company-owned F&B branches across UAE",
    profileSummary: "Supports CCTV and branch infrastructure work including IT asset installation, site support, troubleshooting and project execution.",
  },
  {
    name: "Demo Team Lead",
    username: "demo-lead",
    roleTitle: "IT Team Leader",
    department: "IT Department",
    company: "Demo Restaurant Group",
    reportingTo: "IT Manager",
    specialization: "Team coordination, project supervision, approvals and escalation management",
    locationScope: "Company-owned F&B branches across UAE",
    profileSummary: "Coordinates IT team work, branch opening support, task assignment, review, approvals and escalation handling.",
  },
  {
    name: "Previous Team",
    username: "previous",
    roleTitle: "Legacy / Previous IT Team",
    department: "IT Department",
    company: "Demo Restaurant Group",
    reportingTo: "N/A",
    specialization: "Existing branch setup completed before current records",
    locationScope: "Legacy branches",
    profileSummary: "Used for separating old branch projects completed before Demo Technician joined or before current documentation started.",
  },
];

const defaultGroups = [
  {
    id: 1,
    name: "Demo Restaurant Group",
    concepts: ["Demo Restaurant", "Demo QSR", "Demo Dining", "Demo Brand", "Brand Five", "Demo Beverage", "Demo Cafe", "Demo Concept"],
    logoDataUrl: "",
    footerPhone: "Not configured",
    footerAddress: "Demo Head Office",
    footerWebsite: "example.com",
    footerColor: "#991b1e",
  },
  { id: 2, name: "Demo Premium Dining", concepts: ["Demo Premium Dining"], logoDataUrl: "", footerPhone: "", footerAddress: "", footerWebsite: "", footerColor: "#991b1e" },
  { id: 3, name: "Demo Grill Group", concepts: ["Demo Grill Group", "Demo Grill", "Momo"], logoDataUrl: "", footerPhone: "", footerAddress: "", footerWebsite: "", footerColor: "#991b1e" },
  { id: 4, name: "Demo Bistro", concepts: ["Demo Bistro"], logoDataUrl: "", footerPhone: "", footerAddress: "", footerWebsite: "", footerColor: "#991b1e" },
];

const defaultLocations = [
  { id: 1, name: "Dubai Mall", city: "Dubai", emirate: "Dubai", type: "Mall" },
  { id: 2, name: "Dubai Festival City", city: "Dubai", emirate: "Dubai", type: "Mall" },
  { id: 3, name: "Yas Mall", city: "Abu Dhabi", emirate: "Abu Dhabi", type: "Mall" },
  { id: 4, name: "Reem Mall", city: "Abu Dhabi", emirate: "Abu Dhabi", type: "Mall" },
  { id: 5, name: "Marina Mall", city: "Abu Dhabi", emirate: "Abu Dhabi", type: "Mall" },
  { id: 6, name: "Zahia City Centre", city: "Sharjah", emirate: "Sharjah", type: "Mall" },
];

const defaultBranches = [
  { id: 1, groupName: "Demo Restaurant Group", conceptName: "Demo QSR", branchName: "Demo QSR", locationName: "Zahia City Centre", city: "Sharjah", emirate: "Sharjah", branchType: "QSR", openingDate: "2026-05-23", completedDate: "2026-05-22", status: "Opening Today", branchManager: "To be updated", contactNumber: "To be updated", preparedBy: "Demo Technician", handover: 92, notes: "New QSR branch. IT rack, switch, NVR, cameras, POS, biometric and Etisalat internet setup completed." },
  { id: 2, groupName: "Demo Restaurant Group", conceptName: "Demo Restaurant", branchName: "Chinese Palace Rest.", locationName: "Dubai Mall", city: "Dubai", emirate: "Dubai", branchType: "Dining", openingDate: "Live Branch", completedDate: "2026-05-10", status: "Live", branchManager: "To be updated", contactNumber: "To be updated", preparedBy: "IT Team", handover: 100, notes: "Sample live branch record for preview." },
  { id: 3, groupName: "Demo Restaurant Group", conceptName: "Demo Dining", branchName: "Koryo", locationName: "Dubai Festival City", city: "Dubai", emirate: "Dubai", branchType: "QSR", openingDate: "Live Branch", completedDate: "2026-05-01", status: "Live", branchManager: "To be updated", contactNumber: "To be updated", preparedBy: "IT Team", handover: 85, notes: "Sample Koryo branch record for preview." },
];

const defaultAssets = [
  { id: 101, branchId: 1, category: "NVR", deviceName: "Hikvision NVR", brand: "Hikvision", model: "To be updated", serialNumber: "NVR-UMM-001", ipAddress: "192.168.20.10", macAddress: "Not updated", installedLocation: "IT Rack", switchPort: "SW01-P01", channelNumber: "", cameraView: "", warrantyStart: "2026-05-23", warrantyEnd: "2027-05-22", status: "Installed", remarks: "Recording tested.", nvrUsername: "admin", nvrPassword: "To be updated", recordingDays: 45, lastRecordingReviewDate: "2026-05-23" },
  { id: 102, branchId: 1, category: "CCTV Camera", deviceName: "Camera - Cash Counter", brand: "Hikvision", model: "Dome Camera", serialNumber: "CAM-UMM-001", ipAddress: "192.168.20.101", macAddress: "Not updated", installedLocation: "Cash Counter", switchPort: "SW01-P05", channelNumber: "CH-01", cameraView: "Cash Counter", warrantyStart: "2026-05-23", warrantyEnd: "2027-05-22", status: "Online", remarks: "View clear and recording tested." },
  { id: 103, branchId: 1, category: "Network Switch", deviceName: "24-Port PoE Switch", brand: "To be updated", model: "24-Port PoE", serialNumber: "SW-UMM-001", ipAddress: "192.168.10.2", macAddress: "Not updated", installedLocation: "IT Rack", switchPort: "Rack", channelNumber: "", cameraView: "", warrantyStart: "2026-05-23", warrantyEnd: "2027-05-22", status: "Installed", remarks: "Switch installed in rack." },
  { id: 104, branchId: 1, category: "POS Terminal", deviceName: "POS Counter 01", brand: "To be updated", model: "To be updated", serialNumber: "POS-UMM-001", ipAddress: "192.168.30.20", macAddress: "Not updated", installedLocation: "Cash Counter", switchPort: "SW01-P12", channelNumber: "", cameraView: "", warrantyStart: "2026-05-23", warrantyEnd: "2027-05-22", status: "Tested", remarks: "POS network tested." },
  { id: 201, branchId: 2, category: "NVR", deviceName: "Hikvision NVR", brand: "Hikvision", model: "To be updated", serialNumber: "NVR-CP-DM-001", ipAddress: "192.168.20.10", macAddress: "Not updated", installedLocation: "IT Rack", switchPort: "SW01-P01", channelNumber: "", cameraView: "", warrantyStart: "Not updated", warrantyEnd: "Not updated", status: "Installed", remarks: "Sample device.", nvrUsername: "admin", nvrPassword: "To be updated", recordingDays: 75, lastRecordingReviewDate: "2026-05-20" },
  { id: 301, branchId: 3, category: "CCTV Camera", deviceName: "Camera - Entrance", brand: "Hikvision", model: "Dome Camera", serialNumber: "CAM-KOR-DFC-001", ipAddress: "192.168.20.101", macAddress: "Not updated", installedLocation: "Entrance", switchPort: "SW01-P04", channelNumber: "CH-01", cameraView: "Entrance", warrantyStart: "Not updated", warrantyEnd: "Not updated", status: "Online", remarks: "Sample device." },
];

const defaultNetworkSetups = [
  { branchId: 1, isp: "Etisalat", internetAccount: "To be updated", bandwidth: "To be updated", routerModel: "To be updated", routerIp: "192.168.10.1", switchModel: "24-Port PoE Switch", nvrIp: "192.168.20.10", cctvRange: "192.168.20.100 - 192.168.20.150", posRange: "192.168.30.20 - 192.168.30.50", biometricIp: "192.168.40.10", wifiName: "", wifiPassword: "", vlanDetails: "Optional / To be updated", remarks: "Etisalat ONT connected to router/firewall. Switch, NVR, POS and biometric connected from rack." },
  { branchId: 2, isp: "Etisalat", internetAccount: "To be updated", bandwidth: "To be updated", routerModel: "To be updated", routerIp: "192.168.10.1", switchModel: "To be updated", nvrIp: "192.168.20.10", cctvRange: "192.168.20.100 - 192.168.20.180", posRange: "192.168.30.20 - 192.168.30.80", biometricIp: "192.168.40.10", wifiName: "", wifiPassword: "", vlanDetails: "Optional / To be updated", remarks: "Sample network setup." },
];

const defaultHandovers = [
  { branchId: 1, handoverDate: "2026-05-23", storeManagerName: "To be updated", receivedBy: "Store Manager / Operations", completedPersons: "Demo Technician, Demo Engineer", inchargeName: "Demo Team Lead", inchargeRole: "Team Leader", handoverStatus: "Pending Signature", operationsRemarks: "Branch opening support completed. Final store manager acknowledgement pending.", itRemarks: "IT rack, switch, NVR, CCTV cameras, POS terminal, biometric and network setup completed." },
];

const defaultPullouts = [
  { branchId: 2, pulloutDate: "2026-05-24", closureDate: "2026-05-24", storeManagerName: "To be updated", removedBy: "Demo Technician, Demo Engineer", inchargeName: "Demo Team Lead", inchargeRole: "Team Leader", receivedBy: "IT Store / Warehouse", destination: "IT Store / Warehouse", pulloutReason: "Branch closed / relocation", pulloutStatus: "Pending Signature", assetCondition: "Devices removed and pending final inspection.", operationsRemarks: "Operations confirmation pending.", itRemarks: "IT assets removed from branch and prepared for return/storage." },
];

const defaultFaultLogs = [
  { id: 1, branchId: 1, faultDate: "2026-05-23", faultTime: "10:30", category: "CCTV", title: "NVR recording days below standard", description: "Recording days checked and found below 60 days.", reportedBy: "Daily CCTV review", attendedBy: "Demo Technician", actionTaken: "Checked NVR storage and marked for HDD/storage review.", status: "Open", closedDate: "", remarks: "Follow-up required to increase recording days." },
  { id: 2, branchId: 2, faultDate: "2026-05-20", faultTime: "14:15", category: "Network", title: "POS network intermittent", description: "POS counter network was unstable during operation.", reportedBy: "Store Manager", attendedBy: "Demo Engineer", actionTaken: "Checked switch port and replaced patch cable.", status: "Closed", closedDate: "2026-05-20", remarks: "Issue resolved after cable replacement." },
];

const emptyBranch = { groupName: "Demo Restaurant Group", conceptName: "Demo QSR", branchName: "", locationName: "Zahia City Centre", branchType: "QSR", openingDate: "", completedDate: "", status: "Planning", branchManager: "", contactNumber: "", preparedBy: "Demo Technician", notes: "" };
const emptyLocation = { name: "", city: "", emirate: "Dubai", type: "Mall", notes: "" };
const emptyAsset = { category: "CCTV Camera", customCategory: "", deviceName: "", description: "", quantity: "1", brand: "", model: "", serialNumber: "", ipAddress: "", macAddress: "", installedLocation: "", switchPort: "", channelNumber: "", cameraView: "", warrantyStart: "", warrantyEnd: "", status: "Installed", remarks: "" };
const emptyNetwork = { isp: "Etisalat", internetAccount: "", bandwidth: "", routerModel: "", routerIp: "", switchModel: "", nvrIp: "", cctvRange: "", posRange: "", biometricIp: "", wifiName: "", wifiPassword: "", vlanDetails: "", remarks: "" };
const emptyHandover = { handoverDate: "", storeManagerName: "", receivedBy: "Store Manager / Operations", completedPersons: "Demo Technician, Demo Engineer", inchargeName: "Demo Team Lead", inchargeRole: "Team Leader", handoverStatus: "Pending Signature", operationsRemarks: "", itRemarks: "" };
const emptyPullout = { pulloutDate: "", closureDate: "", storeManagerName: "", removedBy: "Demo Technician, Demo Engineer", inchargeName: "Demo Team Lead", inchargeRole: "Team Leader", receivedBy: "IT Store / Warehouse", destination: "IT Store / Warehouse", pulloutReason: "Branch closed / relocation", pulloutStatus: "Pending Signature", assetCondition: "", operationsRemarks: "", itRemarks: "" };
const emptyFault = { branchId: "", faultDate: "", faultTime: "", category: "CCTV", title: "", description: "", reportedBy: "", attendedBy: "Demo Technician", actionTaken: "", status: "Open", closedDate: "", remarks: "" };
const emptyProfile = { name: "", username: "", password: "", accessLevel: "user", email: "", mobile: "", roleTitle: "", department: "IT Department", company: "Demo Restaurant Group", reportingTo: "", specialization: "", locationScope: "Company-owned F&B branches across UAE", profileSummary: "" };
const emptyGroupBranding = { id: "", name: "", logoDataUrl: "", footerPhone: "", footerAddress: "", footerWebsite: "", footerColor: "#991b1e" };

const emptyCredential = {
  type: "device",
  branchId: "",
  category: "NVR",
  deviceName: "",
  linkedAssetId: "",
  linkedCredentialId: "",
  platform: "",
  accountName: "",
  scopeType: "QSR",
  scopeName: "",
  ipAddress: "",
  url: "",
  port: "",
  username: "",
  email: "",
  password: "",
  recoveryEmail: "",
  mobile: "",
  mfaStatus: "Not Enabled",
  owner: "",
  jobRole: "",
  departmentOperation: "",
  permissionLevel: "Viewer",
  grantedDate: "",
  approvedBy: "",
  revokedDate: "",
  lastPasswordChange: "",
  nextPasswordChange: "",
  status: "Active",
  notes: "",
};

const credentialTypeLabels = {
  shared: "Shared Accounts",
  device: "Device Credentials",
  access: "Access Register",
};

const credentialScopes = ["QSR", "DTF", "Stockroom", "Office", "Warehouse", "Management", "Other"];

function normalizeCredentialRecord(item = {}) {
  const type = item.type === "cloud" ? "shared" : (item.type || "device");
  return {
    ...emptyCredential,
    ...item,
    type,
    scopeType: item.scopeType || item.branchAccess || (type === "shared" ? "Other" : ""),
    scopeName: item.scopeName || "",
    linkedCredentialId: item.linkedCredentialId || "",
    jobRole: item.jobRole || "",
    departmentOperation: item.departmentOperation || item.scopeType || "",
  };
}

const defaultAssetProducts = [
  { id: 1, productName: "NVR - Hikvision DS-7608NI-Q2", category: "NVR", brand: "Hikvision", model: "DS-7608NI-Q2", description: "Hikvision 8 channel NVR", defaultQuantity: "1", defaultLocation: "IT Rack", defaultStatus: "Installed", remarks: "" },
  { id: 2, productName: "NVR - Hikvision DS-7608NXI-K2", category: "NVR", brand: "Hikvision", model: "DS-7608NXI-K2", description: "Hikvision AcuSense NVR", defaultQuantity: "1", defaultLocation: "IT Rack", defaultStatus: "Installed", remarks: "" },
  { id: 3, productName: "Camera - Hikvision Dome", category: "CCTV Camera", brand: "Hikvision", model: "Dome Camera", description: "Indoor CCTV dome camera", defaultQuantity: "1", defaultLocation: "Ceiling", defaultStatus: "Online", remarks: "" },
  { id: 4, productName: "Camera - Hikvision Bullet", category: "CCTV Camera", brand: "Hikvision", model: "Bullet Camera", description: "Outdoor CCTV bullet camera", defaultQuantity: "1", defaultLocation: "Entrance", defaultStatus: "Online", remarks: "" },
  { id: 5, productName: "Network Switch - 24 Port PoE", category: "Network Switch", brand: "", model: "24-Port PoE", description: "24 port PoE network switch", defaultQuantity: "1", defaultLocation: "IT Rack", defaultStatus: "Installed", remarks: "" },
  { id: 6, productName: "Router / Firewall", category: "Router / Firewall", brand: "", model: "", description: "Branch router or firewall", defaultQuantity: "1", defaultLocation: "IT Rack", defaultStatus: "Installed", remarks: "" },
  { id: 7, productName: "POS Terminal", category: "POS Terminal", brand: "", model: "", description: "POS billing terminal", defaultQuantity: "1", defaultLocation: "Cash Counter", defaultStatus: "Installed", remarks: "" },
  { id: 8, productName: "Receipt Printer", category: "Receipt Printer", brand: "", model: "Thermal Printer", description: "POS receipt printer", defaultQuantity: "1", defaultLocation: "Cash Counter", defaultStatus: "Installed", remarks: "" },
  { id: 9, productName: "Kitchen Printer", category: "Receipt Printer", brand: "", model: "Thermal Printer", description: "Kitchen order printer", defaultQuantity: "1", defaultLocation: "Kitchen", defaultStatus: "Installed", remarks: "" },
  { id: 10, productName: "Access Point", category: "Access Point", brand: "", model: "", description: "Wireless access point", defaultQuantity: "1", defaultLocation: "Ceiling", defaultStatus: "Installed", remarks: "" },
  { id: 11, productName: "UPS", category: "UPS", brand: "", model: "", description: "Power backup UPS", defaultQuantity: "1", defaultLocation: "IT Rack", defaultStatus: "Installed", remarks: "" },
  { id: 12, productName: "Biometric Device", category: "Biometric Device", brand: "", model: "", description: "Attendance biometric device", defaultQuantity: "1", defaultLocation: "Staff Entry", defaultStatus: "Installed", remarks: "" },
];

const defaultAssetCategories = [
  "CCTV Camera", "NVR", "Network Switch", "Router / Firewall", "Access Point", "POS Terminal", "Receipt Printer", "KDS Screen", "Biometric Device", "Door Alarm", "UPS", "IT Rack", "Other IT Device"
].map((name, index) => ({ id: index + 1, name, description: "", isActive: true }));

const emptyAssetProduct = { productName: "", category: "CCTV Camera", brand: "", model: "", description: "", defaultQuantity: "1", defaultLocation: "", defaultStatus: "Installed", isActive: true, remarks: "" };

function normalizeAssetProduct(item = {}) {
  return { ...emptyAssetProduct, ...item, isActive: item.isActive !== false };
}

function mergeAssetCategories(categories = [], products = []) {
  const source = Array.isArray(categories) && categories.length ? categories : defaultAssetCategories;
  const next = source.map((item) => ({ ...item, isActive: item.isActive !== false }));
  const names = new Set(next.map((item) => normalizeExcelText(item.name)));
  products.forEach((product) => {
    const name = String(product.category || "").trim();
    if (name && !names.has(normalizeExcelText(name))) {
      next.push({ id: `category-${normalizeExcelText(name)}-${next.length}`, name, description: "Recovered from existing product data", isActive: true });
      names.add(normalizeExcelText(name));
    }
  });
  return next.sort((a, b) => a.name.localeCompare(b.name));
}

defaultAssetProducts.forEach((item) => { if (typeof item.isActive !== "boolean") item.isActive = true; });

const productMasterFields = [
  ["productName", "Product Name", true],
  ["category", "Device Category", true],
  ["brand", "Brand", false],
  ["model", "Model", false],
  ["description", "Description", false],
  ["defaultQuantity", "Default Quantity", false],
  ["defaultLocation", "Default Installed Location", false],
  ["defaultStatus", "Default Status", false],
  ["remarks", "Remarks", false],
];


const emptyProductImport = {
  fileName: "",
  sheetNames: [],
  selectedSheet: "",
  headers: [],
  rows: [],
  mapping: {
    productName: "",
    category: "",
    brand: "",
    model: "",
    description: "",
    defaultQuantity: "",
    defaultLocation: "",
    defaultStatus: "",
    remarks: "",
  },
  duplicateMode: "skip",
  preview: null,
};

const productColumnAliases = {
  productName: ["product name", "asset product", "item name", "device name", "product", "name"],
  category: ["device category", "product category", "asset category", "category", "type"],
  brand: ["brand", "manufacturer", "make", "vendor"],
  model: ["model", "model no", "model number", "part number"],
  description: ["description", "product description", "details", "specification", "specifications"],
  defaultQuantity: ["default quantity", "quantity", "qty", "default qty"],
  defaultLocation: ["default installed location", "default location", "installed location", "installation location", "location"],
  defaultStatus: ["default status", "status", "asset status"],
  remarks: ["remarks", "remark", "notes", "comments"],
};

const emptyTemplateRow = { productId: "", quantity: "1", installedLocation: "", remarks: "" };
const emptyAssetTemplate = { name: "", branchType: "QSR", description: "", rows: [{ ...emptyTemplateRow }] };

const defaultAssetTemplates = [
  {
    id: 1,
    name: "Standard QSR Branch Setup",
    branchType: "QSR",
    description: "Default asset setup for normal QSR / mall branch opening.",
    rows: [
      { productId: 7, quantity: "2", installedLocation: "Cash Counter", remarks: "Update serial number after installation" },
      { productId: 6, quantity: "1", installedLocation: "IT Rack", remarks: "Router/firewall for branch internet" },
      { productId: 5, quantity: "1", installedLocation: "IT Rack", remarks: "Main branch switch" },
      { productId: 1, quantity: "1", installedLocation: "IT Rack", remarks: "NVR recording to be reviewed" },
      { productId: 3, quantity: "8", installedLocation: "Store Area", remarks: "Camera quantity can be adjusted" },
      { productId: 9, quantity: "2", installedLocation: "Kitchen", remarks: "KOT printer" },
      { productId: 10, quantity: "2", installedLocation: "Ceiling", remarks: "Wi-Fi coverage" },
      { productId: 11, quantity: "1", installedLocation: "IT Rack", remarks: "Power backup" },
    ],
  },
  {
    id: 2,
    name: "Accommodation CCTV Setup",
    branchType: "Accommodation",
    description: "Basic CCTV/NVR setup for accommodation or staff facility.",
    rows: [
      { productId: 1, quantity: "1", installedLocation: "Security / Network Area", remarks: "NVR" },
      { productId: 3, quantity: "16", installedLocation: "Common Area", remarks: "Adjust camera quantity as per site" },
      { productId: 5, quantity: "1", installedLocation: "Network Rack", remarks: "PoE switch" },
      { productId: 11, quantity: "1", installedLocation: "Network Rack", remarks: "UPS" },
    ],
  },
];

const assetTemplateFields = [
  ["templateName", "Template Name", true],
  ["branchType", "Branch Type", false],
  ["productName", "Product Name", true],
  ["category", "Device Category", false],
  ["brand", "Brand", false],
  ["model", "Model", false],
  ["quantity", "Quantity", true],
  ["installedLocation", "Installed Location", false],
  ["remarks", "Remarks", false],
];


const emptyExcelImport = {
  fileName: "",
  sheetNames: [],
  selectedSheet: "",
  headers: [],
  rows: [],
  mapping: {
    serialNumber: "",
    category: "",
    deviceName: "",
    description: "",
    brand: "",
    model: "",
    quantity: "",
    ipAddress: "",
    macAddress: "",
    installedLocation: "",
    switchPort: "",
    channelNumber: "",
    cameraView: "",
    warrantyStart: "",
    warrantyEnd: "",
    status: "",
    remarks: "",
    nvrUsername: "",
    nvrPassword: "",
    recordingDays: "",
    recordingStartDate: "",
    recordingEndDate: "",
    lastRecordingReviewDate: "",
    reviewedBy: "",
    reviewStatus: "",
    issueFound: "",
    actionRequired: "",
    nextReviewDate: "",
  },
  duplicateMode: "skip",
  preview: null,
};

const excelImportFields = [
  ["serialNumber", "Serial Number", true],
  ["category", "Device Category", true],
  ["deviceName", "Device Name / Asset Name", true],
  ["description", "Description", false],
  ["quantity", "Quantity", false],
  ["brand", "Brand / Manufacturer", false],
  ["model", "Model", false],
  ["ipAddress", "IP Address", false],
  ["macAddress", "MAC Address", false],
  ["installedLocation", "Installed Location", false],
  ["switchPort", "Switch Port", false],
  ["channelNumber", "Camera Channel", false],
  ["cameraView", "Camera View / Area", false],
  ["warrantyStart", "Warranty Start", false],
  ["warrantyEnd", "Warranty End", false],
  ["status", "Status", false],
  ["remarks", "Remarks", false],
  ["nvrUsername", "NVR Username", false],
  ["nvrPassword", "NVR Password", false],
  ["recordingDays", "Recording Days", false],
  ["recordingStartDate", "Recording Start Date", false],
  ["recordingEndDate", "Recording End Date", false],
  ["lastRecordingReviewDate", "Last Review Date", false],
  ["reviewedBy", "Reviewed By", false],
  ["reviewStatus", "Review Status", false],
  ["issueFound", "Issue Found", false],
  ["actionRequired", "Action Required", false],
  ["nextReviewDate", "Next Review Date", false],
];

const excelColumnAliases = {
  serialNumber: ["serial number", "serial no", "serial", "s/n", "sn", "asset serial"],
  category: ["device category", "product name catogery", "product name category", "product name (catogery)", "product name (category)", "category", "catogery", "product category", "product name"],
  deviceName: ["device name / asset name", "asset name", "device name", "item name", "asset", "device", "name"],
  description: ["description", "item description", "asset description", "details", "specification", "specifications"],
  brand: ["manufacturer", "brand", "brand / manufacturer", "make", "vendor"],
  model: ["model", "model no", "model number", "part number", "item model"],
  quantity: ["quantity", "qty", "qnty", "count"],
  ipAddress: ["ip address", "ip", "device ip", "nvr ip"],
  macAddress: ["mac address", "mac", "mac id"],
  installedLocation: ["installed location", "installation location", "asset location", "installed area", "area", "location"],
  switchPort: ["switch port", "port", "switch port number"],
  channelNumber: ["camera channel", "channel", "channel number", "ch"],
  cameraView: ["camera view / area", "camera view", "view area", "camera area"],
  warrantyStart: ["warranty start", "warranty start date"],
  warrantyEnd: ["warranty end", "warranty end date", "warranty expiry"],
  status: ["status", "asset status"],
  remarks: ["remarks", "remark", "notes", "comments"],
  nvrUsername: ["nvr username", "username", "admin username"],
  nvrPassword: ["nvr password", "password", "admin password"],
  recordingDays: ["recording days", "record days", "days"],
  recordingStartDate: ["recording start date", "recording start", "start date"],
  recordingEndDate: ["recording end date", "recording end", "end date"],
  lastRecordingReviewDate: ["last review date", "last recording review date", "review date"],
  reviewedBy: ["reviewed by", "checked by"],
  reviewStatus: ["review status", "nvr status"],
  issueFound: ["issue found", "issue"],
  actionRequired: ["action required", "action", "follow up"],
  nextReviewDate: ["next review date", "next review"],
};


const emptyBranchImport = {
  fileName: "",
  sheetNames: [],
  selectedSheet: "",
  headers: [],
  rows: [],
  mapping: {
    groupName: "",
    conceptName: "",
    branchName: "",
    locationName: "",
    city: "",
    emirate: "",
    branchType: "",
    openingDate: "",
    completedDate: "",
    status: "",
    branchManager: "",
    contactNumber: "",
    preparedBy: "",
    notes: "",
  },
  duplicateMode: "skip",
  preview: null,
};

const branchImportFields = [
  ["groupName", "Main Group", true],
  ["conceptName", "Concept / Restaurant", false],
  ["branchName", "Branch Display Name", true],
  ["locationName", "Location / Mall", true],
  ["city", "City", false],
  ["emirate", "Emirate", false],
  ["branchType", "Branch Type", true],
  ["openingDate", "Opening Date", false],
  ["completedDate", "IT Setup Completed Date", false],
  ["status", "Current Status", false],
  ["branchManager", "Branch Manager", false],
  ["contactNumber", "Contact Number", false],
  ["preparedBy", "Prepared By", false],
  ["notes", "Notes", false],
];

const branchColumnAliases = {
  groupName: ["main group", "group", "company group", "brand group", "business group"],
  conceptName: ["concept", "concept restaurant", "concept / restaurant", "restaurant", "brand", "sub brand"],
  branchName: ["branch display name", "branch name", "display name", "store name", "branch", "name"],
  locationName: ["location mall", "location / mall", "mall", "location", "site", "site name"],
  city: ["city", "area city"],
  emirate: ["emirate", "state", "region"],
  branchType: ["branch type", "type", "store type", "format"],
  openingDate: ["opening date", "open date", "store opening date"],
  completedDate: ["it setup completed date", "completed date", "setup completed date", "handover date"],
  status: ["current status", "status", "project status"],
  branchManager: ["branch manager", "store manager", "manager"],
  contactNumber: ["contact number", "phone", "mobile", "telephone"],
  preparedBy: ["prepared by", "created by", "project by", "completed by"],
  notes: ["notes", "remarks", "comments"],
};

const emptyLocationImport = {
  fileName: "",
  sheetNames: [],
  selectedSheet: "",
  headers: [],
  rows: [],
  mapping: {
    name: "",
    city: "",
    emirate: "",
    type: "",
    notes: "",
  },
  duplicateMode: "skip",
  preview: null,
};

const locationImportFields = [
  ["name", "Location / Mall", true],
  ["city", "City", true],
  ["emirate", "Emirate", true],
  ["type", "Location Type", false],
  ["notes", "Notes", false],
];

const locationColumnAliases = {
  name: ["location mall", "location / mall", "mall", "location", "site", "site name", "location name", "branch location"],
  city: ["city", "area city", "town"],
  emirate: ["emirate", "state", "region"],
  type: ["location type", "type", "site type", "mall type"],
  notes: ["notes", "remarks", "comments", "description"],
};

const templateDefinitions = {
  branch: {
    title: "Branch Import Template",
    fileName: "Branch_Import_Template",
    sheetName: "Branches",
    fields: branchImportFields.map(([field, label, required]) => ({ field, label, important: required || ["conceptName", "city", "emirate", "status"].includes(field) })),
    sample: {
      groupName: "Demo Restaurant Group",
      conceptName: "Demo QSR",
      branchName: "Umami Sample Branch",
      locationName: "Dubai Mall",
      city: "Dubai",
      emirate: "Dubai",
      branchType: "Dining",
      openingDate: "2026-07-01",
      completedDate: "2026-07-02",
      status: "Live",
      branchManager: "Manager Name",
      contactNumber: "+971500000000",
      preparedBy: "Demo Technician",
      notes: "Sample row - replace with actual branch details",
    },
  },
  location: {
    title: "Location Import Template",
    fileName: "Location_Import_Template",
    sheetName: "Locations",
    fields: locationImportFields.map(([field, label, required]) => ({ field, label, important: required || field === "type" })),
    sample: { name: "Dubai Mall", city: "Dubai", emirate: "Dubai", type: "Mall", notes: "Sample row - replace with actual location" },
  },
  asset: {
    title: "IT Asset Import Template",
    fileName: "IT_Asset_Import_Template",
    sheetName: "IT Assets",
    fields: excelImportFields.map(([field, label, required]) => ({ field, label, important: required || ["description", "quantity", "brand", "model", "installedLocation", "status"].includes(field) })),
    sample: {
      serialNumber: "SN-SAMPLE-001",
      category: "NVR",
      deviceName: "Hikvision NVR",
      description: "8/16 Channel NVR",
      quantity: "1",
      brand: "Hikvision",
      model: "DS-7608NI-Q2",
      ipAddress: "192.168.20.10",
      macAddress: "",
      installedLocation: "IT Rack",
      switchPort: "SW01-P01",
      channelNumber: "",
      cameraView: "",
      warrantyStart: "2026-07-01",
      warrantyEnd: "2027-07-01",
      status: "Installed",
      remarks: "Sample row - replace with actual asset details",
      nvrUsername: "admin",
      nvrPassword: "",
      recordingDays: "60",
      recordingStartDate: "2026-05-01",
      recordingEndDate: "2026-06-29",
      lastRecordingReviewDate: "2026-06-29",
      reviewedBy: "Demo Technician",
      reviewStatus: "OK",
      issueFound: "",
      actionRequired: "",
      nextReviewDate: "2026-07-06",
    },
  },
  product: {
    title: "Asset Product List Template",
    fileName: "Asset_Product_List_Template",
    sheetName: "Product List",
    fields: productMasterFields.map(([field, label, required]) => ({ field, label, important: required || ["brand", "model", "description", "defaultQuantity", "defaultLocation", "defaultStatus"].includes(field) })),
    sample: {
      productName: "NVR - Hikvision DS-7608NI-Q2",
      category: "NVR",
      brand: "Hikvision",
      model: "DS-7608NI-Q2",
      description: "Hikvision 8 channel NVR",
      defaultQuantity: "1",
      defaultLocation: "IT Rack",
      defaultStatus: "Installed",
      remarks: "Sample row - replace or add your product list",
    },
  },
  branchAssetTemplate: {
    title: "Branch Asset Template Builder Excel",
    fileName: "Branch_Asset_Template_Builder",
    sheetName: "Branch Asset Template",
    fields: assetTemplateFields.map(([field, label, required]) => ({ field, label, important: required || ["category", "brand", "model", "installedLocation"].includes(field) })),
    sample: {
      templateName: "Standard QSR Branch Setup",
      branchType: "QSR",
      productName: "POS Terminal",
      category: "POS Terminal",
      brand: "Sunmi",
      model: "T2 Mini",
      quantity: "2",
      installedLocation: "Cash Counter",
      remarks: "Sample row - add more rows for switch, router, NVR, camera, printer, UPS",
    },
  },
};

function defaultTemplateSelected(type) {
  const def = templateDefinitions[type] || templateDefinitions.asset;
  return Object.fromEntries(def.fields.map((field) => [field.field, !!field.important]));
}

function detectBranchMapping(headers) {
  const normalized = headers.map((header) => ({ original: header, normalized: normalizeExcelText(header) }));
  const mapping = { ...emptyBranchImport.mapping };
  Object.entries(branchColumnAliases).forEach(([field, aliases]) => {
    const aliasSet = aliases.map(normalizeExcelText);
    const exact = normalized.find((header) => aliasSet.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const partial = normalized.find((header) => aliasSet.some((alias) => header.normalized.includes(alias) || alias.includes(header.normalized)));
    if (partial) mapping[field] = partial.original;
  });
  return mapping;
}

function normalizeBranchKey(branchName, locationName) {
  return `${normalizeExcelText(branchName)}|${normalizeExcelText(locationName)}`;
}
function detectLocationMapping(headers) {
  const normalized = headers.map((header) => ({ original: header, normalized: normalizeExcelText(header) }));
  const mapping = { ...emptyLocationImport.mapping };
  Object.entries(locationColumnAliases).forEach(([field, aliases]) => {
    const aliasSet = aliases.map(normalizeExcelText);
    const exact = normalized.find((header) => aliasSet.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const partial = normalized.find((header) => aliasSet.some((alias) => header.normalized.includes(alias) || alias.includes(header.normalized)));
    if (partial) mapping[field] = partial.original;
  });
  return mapping;
}

function normalizeLocationKey(name, city, emirate) {
  return `${normalizeExcelText(name)}|${normalizeExcelText(city)}|${normalizeExcelText(emirate)}`;
}

function cleanLocationType(value) {
  const raw = String(value || "").trim();
  const allowed = ["Mall", "Street", "Airport", "Office", "Warehouse", "Cloud Kitchen", "Other"];
  const found = allowed.find((item) => normalizeExcelText(item) === normalizeExcelText(raw));
  return found || raw || "Mall";
}


function cleanBranchStatus(value) {
  const raw = String(value || "").trim();
  const allowed = ["Planning", "Installation Started", "Pending Handover", "Opening Today", "Live", "Closed"];
  const found = allowed.find((item) => normalizeExcelText(item) === normalizeExcelText(raw));
  return found || "Planning";
}

function cleanBranchType(value) {
  const raw = String(value || "").trim();
  const allowed = ["QSR", "Dining", "Cloud Kitchen", "Kiosk", "Office", "Warehouse"];
  const found = allowed.find((item) => normalizeExcelText(item) === normalizeExcelText(raw));
  return found || raw || "QSR";
}

function cleanDateText(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return raw;
}

function normalizeExcelText(value) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function normalizeSerial(value) {
  return String(value ?? "").trim().toLowerCase();
}

function detectExcelMapping(headers) {
  const normalized = headers.map((header) => ({ original: header, normalized: normalizeExcelText(header) }));
  const mapping = { ...emptyExcelImport.mapping };
  Object.entries(excelColumnAliases).forEach(([field, aliases]) => {
    const aliasSet = aliases.map(normalizeExcelText);
    const exact = normalized.find((header) => aliasSet.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const partial = normalized.find((header) => aliasSet.some((alias) => header.normalized.includes(alias) || alias.includes(header.normalized)));
    if (partial) mapping[field] = partial.original;
  });
  return mapping;
}


function detectProductMapping(headers) {
  const normalized = headers.map((header) => ({ original: header, normalized: normalizeExcelText(header) }));
  const mapping = { ...emptyProductImport.mapping };
  Object.entries(productColumnAliases).forEach(([field, aliases]) => {
    const aliasSet = aliases.map(normalizeExcelText);
    const exact = normalized.find((header) => aliasSet.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const partial = normalized.find((header) => aliasSet.some((alias) => header.normalized.includes(alias) || alias.includes(header.normalized)));
    if (partial) mapping[field] = partial.original;
  });
  return mapping;
}

function normalizeProductKey(product = {}) {
  const category = normalizeExcelText(product.category);
  const brand = normalizeExcelText(product.brand);
  const model = normalizeExcelText(product.model);
  if (category && (brand || model)) return `model|${category}|${brand}|${model}`;
  return `name|${normalizeExcelText(product.productName)}`;
}

function excelValue(row, column) {
  if (!column) return "";
  return String(row?.[column] ?? "").trim();
}

function mapExcelCategory(category) {
  const raw = String(category || "").trim();
  const key = normalizeExcelText(raw);
  if (!raw) return "Other IT Device";
  if (key.includes("camera")) return "CCTV Camera";
  if (key === "nvr" || key.includes("nvr")) return "NVR";
  if (key.includes("network switch") || key.includes("switch")) return "Network Switch";
  if (key.includes("router") || key.includes("firewall")) return "Router / Firewall";
  if (key === "pos" || key.includes("point of sale")) return "POS Terminal";
  if (key.includes("receipt") || key.includes("thermal printer") || key.includes("printer")) return "Receipt Printer";
  if (key.includes("biometric")) return "Biometric Device";
  if (key.includes("door alarm")) return "Door Alarm";
  if (key.includes("ups")) return "UPS";
  if (key.includes("access point") || key === "ap") return "Access Point";
  if (key.includes("rack")) return "IT Rack";
  return raw;
}

function cleanExcelQuantity(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "1";
  const num = Number(raw);
  if (Number.isFinite(num) && num > 0) return String(num);
  return raw;
}

function cleanImportedPlaceholder(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const normalized = normalizeExcelText(raw);
  if (normalized === "import from excel sheet" || normalized === "imported from excel sheet" || normalized.startsWith("imported from excel")) return "";
  return raw;
}

function cleanReportValue(value, fallback = "") {
  const raw = cleanImportedPlaceholder(value);
  const normalized = normalizeExcelText(raw);
  if (!raw || normalized === "to be updated" || normalized === "not updated" || normalized === "not assigned" || normalized === "no remarks" || normalized === "no remarks updated") return fallback;
  return raw;
}

const LS = {
  users: "branch_it_dashboard_users_v3",
  profiles: "branch_it_dashboard_user_profiles_v3",
  groups: "branch_it_dashboard_groups_v3",
  locations: "branch_it_dashboard_locations_v3",
  branches: "branch_it_dashboard_branches_v3",
  assets: "branch_it_dashboard_assets_v3",
  networks: "branch_it_dashboard_networks_v3",
  handovers: "branch_it_dashboard_handovers_v3",
  pullouts: "branch_it_dashboard_pullouts_v3",
  faults: "branch_it_dashboard_fault_logs_v3",
  assetProducts: "branch_it_dashboard_asset_products_v3",
  assetCategories: "branch_it_dashboard_asset_categories_v1",
  assetTemplates: "branch_it_dashboard_asset_templates_v3",
  credentials: "branch_it_dashboard_credentials_v1",
  vaultSettings: "branch_it_dashboard_vault_settings_v1",
  resetRequests: "branch_it_dashboard_password_reset_requests_v1",
  theme: "branch_it_dashboard_theme_v3",
};

function load(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    if (!value) return fallback;
    const parsed = JSON.parse(value);
    return Array.isArray(fallback) ? (Array.isArray(parsed) ? parsed : fallback) : parsed;
  } catch {
    return fallback;
  }
}
function save(key, value) { localStorage.setItem(key, JSON.stringify(value)); }
function today() { return new Date().toISOString().slice(0, 10); }
function toUtcDateValue(value) {
  if (!value) return null;
  const parts = String(value).split("-").map(Number);
  if (parts.length !== 3 || parts.some((part) => Number.isNaN(part))) return null;
  return Date.UTC(parts[0], parts[1] - 1, parts[2]);
}
function countRecordingDays(startDate, endDate) {
  const start = toUtcDateValue(startDate);
  const end = toUtcDateValue(endDate);
  if (start === null || end === null || end < start) return 0;
  return Math.floor((end - start) / (1000 * 60 * 60 * 24)) + 1;
}
function addDays(dateString, days) {
  const base = toUtcDateValue(dateString || today());
  const date = new Date((base === null ? toUtcDateValue(today()) : base) + days * 24 * 60 * 60 * 1000);
  return date.toISOString().slice(0, 10);
}
function getAutoNvrStatus(days) {
  const n = Number(days || 0);
  if (!n) return "Review Pending";
  return n < 60 ? "Below 60 Days" : "OK";
}
function makeId() { return Date.now() + Math.floor(Math.random() * 999); }

function bytesToBase64(bytes) {
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary);
}
function base64ToBytes(value) {
  const binary = atob(value);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}
async function deriveVaultKey(passphrase, saltBase64) {
  const material = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(passphrase),
    "PBKDF2",
    false,
    ["deriveKey"]
  );
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", salt: base64ToBytes(saltBase64), iterations: 250000, hash: "SHA-256" },
    material,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
}
async function encryptVaultText(value, key) {
  if (!value) return "";
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const cipher = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    key,
    new TextEncoder().encode(value)
  );
  return { version: 1, iv: bytesToBase64(iv), cipher: bytesToBase64(new Uint8Array(cipher)) };
}
async function decryptVaultText(payload, key) {
  if (!payload?.iv || !payload?.cipher) return "";
  const plain = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: base64ToBytes(payload.iv) },
    key,
    base64ToBytes(payload.cipher)
  );
  return new TextDecoder().decode(plain);
}

function downloadCredentialTemplate() {
  const workbook = XLSX.utils.book_new();
  const sheets = [
    {
      name: "Shared Accounts",
      rows: [{
        Platform: "Hik-Connect", "Account Scope": "QSR", "Scope Name": "QSR Supervisors",
        "Account Name": "Hik-Connect - QSR Supervisors", "Login Email": "qsr@example.com",
        Username: "qsr-supervisor", "Current Password": "sample-password", "Account Owner": "IT Department",
        "MFA Status": "Not Enabled", Status: "Active", "Last Password Changed": "2026-07-01",
        "Next Review Date": "2026-10-01", "Recovery Email": "", "Recovery Mobile": "",
        Notes: "Replace this sample row before import"
      }],
    },
    {
      name: "Device Credentials",
      rows: [{
        Branch: "HO - JLT", "Linked Asset Serial": "NVR-001", "Device Type": "NVR",
        "Device Name": "Main CCTV NVR", "IP Address": "192.168.20.10", URL: "", Port: "8000",
        Username: "admin", Password: "sample-password", Status: "Active",
        "Last Password Change": "2026-07-01", "Next Review Date": "2026-10-01",
        Notes: "Replace this sample row before import"
      }],
    },
    {
      name: "User Access Register",
      rows: [{
        "Linked Account": "Hik-Connect - QSR Supervisors", "User Full Name": "Example User",
        "Job Role": "Supervisor", "Department / Operation": "QSR", Email: "user@example.com",
        Mobile: "", "Permission Level": "Viewer", "Granted Date": "2026-07-01",
        "Approved By": "IT Manager", Status: "Active", "Revoked Date": "",
        Notes: "Replace this sample row before import"
      }],
    },
  ];
  sheets.forEach(({ name, rows }) => {
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = Object.keys(rows[0]).map((header) => ({ wch: Math.max(18, header.length + 3) }));
    XLSX.utils.book_append_sheet(workbook, worksheet, name);
  });
  XLSX.writeFile(workbook, `Credential_Vault_Import_Template_${today()}.xlsx`);
}

function Badge({ children, type = "default" }) {
  const map = {
    default: "bg-slate-100 text-slate-700",
    success: "bg-green-100 text-green-700",
    warning: "bg-yellow-100 text-yellow-800",
    danger: "bg-red-100 text-red-700",
    info: "bg-blue-100 text-blue-700",
    purple: "bg-purple-100 text-purple-700",
  };
  return <span className={`inline-flex rounded-full px-3 py-1 text-xs font-semibold ${map[type] || map.default}`}>{children}</span>;
}

function AppIcon({ name = "dashboard", className = "h-5 w-5" }) {
  const paths = {
    dashboard: <><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></>,
    user: <><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></>,
    groups: <><circle cx="9" cy="8" r="3"/><circle cx="17" cy="9" r="2.5"/><path d="M3 20a6 6 0 0 1 12 0M14 15a5 5 0 0 1 7 5"/></>,
    location: <><path d="M20 10c0 5-8 11-8 11S4 15 4 10a8 8 0 1 1 16 0Z"/><circle cx="12" cy="10" r="2.5"/></>,
    branches: <><path d="M4 21V7l8-4 8 4v14"/><path d="M9 21v-5h6v5M8 9h.01M12 9h.01M16 9h.01M8 13h.01M12 13h.01M16 13h.01"/></>,
    assets: <><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 21h8M12 16v5"/></>,
    cctv: <><path d="m4 7 11-3 2 7-11 3z"/><path d="m15.5 8.5 5.5 2.5-2 4.5-5.3-2.4M7 14l-1 5h9"/></>,
    faults: <><path d="M12 3 2.8 19h18.4L12 3Z"/><path d="M12 9v4M12 17h.01"/></>,
    handover: <><path d="M8 12h8M12 8l4 4-4 4"/><path d="M4 5h6v14H4zM14 5h6v14h-6"/></>,
    vault: <><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3M12 14v3"/></>,
    eye: <><path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z"/><circle cx="12" cy="12" r="2.5"/></>,
    eyeOff: <><path d="m3 3 18 18"/><path d="M10.6 10.7a2.5 2.5 0 0 0 3.4 3.4"/><path d="M9.9 5.2A11.4 11.4 0 0 1 12 5c6.5 0 10 7 10 7a16 16 0 0 1-3 3.8M6.6 6.6C3.5 8.4 2 12 2 12s3.5 7 10 7a10.8 10.8 0 0 0 4.1-.8"/></>,
    copy: <><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M15 9V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v7a2 2 0 0 0 2 2h3"/></>,
    lock: <><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></>,
    settings: <><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .34 1.88l.06.06-2.83 2.83-.06-.06A1.7 1.7 0 0 0 15 19.4a1.7 1.7 0 0 0-1 .6 1.7 1.7 0 0 0-.4 1V21h-4v-.09a1.7 1.7 0 0 0-1.1-1.51 1.7 1.7 0 0 0-1.88.35l-.06.06-2.83-2.83.06-.06A1.7 1.7 0 0 0 4.6 15a1.7 1.7 0 0 0-.6-1 1.7 1.7 0 0 0-1-.4H3v-4h.09A1.7 1.7 0 0 0 4.6 8.5a1.7 1.7 0 0 0-.34-1.88l-.06-.06 2.83-2.83.06.06A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-.6 1.7 1.7 0 0 0 .4-1V3h4v.09A1.7 1.7 0 0 0 15.5 4.6a1.7 1.7 0 0 0 1.88-.34l.06-.06 2.83 2.83-.06.06A1.7 1.7 0 0 0 19.4 9c.12.36.34.7.6 1 .28.26.64.4 1 .4H21v4h-.09A1.7 1.7 0 0 0 19.4 15Z"/></>,
    check: <><circle cx="12" cy="12" r="9"/><path d="m8 12 2.5 2.5L16 9"/></>,
    briefcase: <><rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8 7V4h8v3M3 12h18"/></>,
    calendar: <><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></>,
    download: <><path d="M12 3v12M7 10l5 5 5-5"/><path d="M5 21h14"/></>,
    upload: <><path d="M12 21V9M7 14l5-5 5 5"/><path d="M5 3h14"/></>,
    plus: <><circle cx="12" cy="12" r="9"/><path d="M12 8v8M8 12h8"/></>,
    logout: <><path d="M10 17l5-5-5-5M15 12H3"/><path d="M14 3h7v18h-7"/></>,
    sync: <><path d="M20 7h-5V2"/><path d="M20 7a8 8 0 1 0 1 7"/></>,
  };
  return <svg viewBox="0 0 24 24" className={className} fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name] || paths.dashboard}</svg>;
}

function DonutChart({ segments = [], total = 0, label = "Total", sizeClass = "h-32 w-32" }) {
  const radius = 46;
  const circumference = 2 * Math.PI * radius;
  const positiveSegments = segments.filter((segment) => Number(segment.value || 0) > 0);
  const segmentTotal = positiveSegments.reduce((sum, segment) => sum + Number(segment.value || 0), 0);
  let consumed = 0;
  return <div className={`relative shrink-0 ${sizeClass}`}>
    <svg viewBox="0 0 120 120" className="h-full w-full -rotate-90" role="img" aria-label={`${label}: ${total}`}>
      <circle className="donut-track" cx="60" cy="60" r={radius} fill="none" stroke="#e2e8f0" strokeWidth="14" />
      {positiveSegments.map((segment, index) => {
        const length = segmentTotal ? (Number(segment.value || 0) / segmentTotal) * circumference : 0;
        const dashOffset = -consumed;
        consumed += length;
        return <circle key={`${segment.label || "segment"}-${index}`} cx="60" cy="60" r={radius} fill="none" stroke={segment.color} strokeWidth="14" strokeDasharray={`${length} ${Math.max(0, circumference - length)}`} strokeDashoffset={dashOffset} strokeLinecap="butt" />;
      })}
    </svg>
    <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center px-4 text-center">
      <span className="text-2xl font-bold leading-none text-slate-950">{total}</span>
      <span className="mt-1 max-w-[76px] text-[9px] font-medium uppercase leading-3 tracking-wide text-slate-500">{label}</span>
    </div>
  </div>;
}

function Card({ title, value, note, icon, accent = "blue", onClick, active = false }) {
  const iconByTitle = {
    "Completed Projects": "check", "Ongoing Projects": "briefcase", "Upcoming Projects": "calendar",
    Locations: "location", Branches: "branches", "IT Assets": "assets", "Pending Handover": "handover",
    "Total NVR": "cctv", "Total Cameras": "cctv", "QSR Cameras": "cctv", "DTF Cameras": "cctv", "Office Cameras": "cctv",
  };
  const accentClasses = {
    blue: "bg-blue-50 text-blue-600 ring-blue-100", green: "bg-emerald-50 text-emerald-600 ring-emerald-100",
    amber: "bg-amber-50 text-amber-600 ring-amber-100", purple: "bg-violet-50 text-violet-600 ring-violet-100",
    teal: "bg-cyan-50 text-cyan-600 ring-cyan-100", rose: "bg-rose-50 text-rose-600 ring-rose-100",
  };
  const Comp = onClick ? "button" : "div";
  return <Comp onClick={onClick} aria-pressed={onClick ? active : undefined} className={`stat-card group w-full rounded-2xl border bg-white p-4 text-left shadow-[0_8px_28px_rgba(15,23,42,0.06)] transition ${active ? "stat-card-active border-blue-500 ring-2 ring-blue-100" : "border-slate-200/80"} ${onClick ? "hover:-translate-y-0.5 hover:border-blue-300 hover:shadow-[0_14px_34px_rgba(37,99,235,0.10)]" : ""}`}><div className="flex items-start justify-between gap-3"><div data-accent={accent} className={`stat-card-icon flex h-9 w-9 shrink-0 items-center justify-center rounded-xl ring-1 ${accentClasses[accent] || accentClasses.blue}`}><AppIcon name={icon || iconByTitle[title] || "dashboard"} className="h-4 w-4" /></div>{onClick && <span className="text-[10px] font-semibold text-slate-400">{active ? "Selected" : "View details"}</span>}</div><p className="mt-3 text-xs font-semibold text-slate-700">{title}</p><h3 className="mt-1 break-words text-2xl font-bold leading-tight text-slate-950">{value}</h3><p className="mt-1.5 text-[11px] leading-4 text-slate-500">{note}</p></Comp>;
}
function Field({ label, required, children }) {
  return <label className="block"><span className="text-sm font-semibold text-slate-700">{label} {required && <span className="text-red-600">*</span>}</span><div className="mt-1">{children}</div></label>;
}
function TextInput({ value, onChange, placeholder = "", type = "text", readOnly = false }) {
  return <input type={type} value={value || ""} readOnly={readOnly} onChange={(e) => !readOnly && onChange(e.target.value)} placeholder={placeholder} className={`w-full rounded-xl border border-slate-300 px-4 py-2.5 text-sm outline-none focus:border-slate-900 ${readOnly ? "bg-slate-100 text-slate-700" : ""}`} />;
}
function PasswordInput({ value, onChange, placeholder = "", readOnly = false }) {
  const [visible, setVisible] = useState(false);
  return <div className="relative"><input type={visible ? "text" : "password"} value={value || ""} readOnly={readOnly} onChange={(e) => !readOnly && onChange(e.target.value)} placeholder={placeholder} className={`w-full rounded-xl border border-slate-300 px-4 py-2.5 pr-11 text-sm outline-none focus:border-slate-900 ${readOnly ? "bg-slate-100 text-slate-700" : ""}`} /><button type="button" onClick={() => setVisible((current) => !current)} className="absolute inset-y-0 right-0 flex w-11 items-center justify-center text-slate-500 hover:text-slate-900" title={visible ? "Hide password" : "Show password"} aria-label={visible ? "Hide password" : "Show password"}><AppIcon name={visible ? "eyeOff" : "eye"} className="h-4 w-4" /></button></div>;
}
function SelectInput({ value, onChange, children }) {
  return <select value={value || ""} onChange={(e) => onChange(e.target.value)} className="w-full rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm outline-none focus:border-slate-900">{children}</select>;
}
function TextArea({ value, onChange, placeholder = "" }) {
  return <textarea value={value || ""} onChange={(e) => onChange(e.target.value)} placeholder={placeholder} className="mt-1 min-h-24 w-full rounded-xl border border-slate-300 px-4 py-2.5 text-sm outline-none focus:border-slate-900" />;
}

function Table({ headers, rows, renderDesktop, renderMobile }) {
  return <>
    <div className="hidden w-full max-w-full overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-sm md:block">
      <table className="pdf-table min-w-[760px] w-full text-left text-sm">
        <thead className="bg-slate-100 text-slate-600"><tr>{headers.map((h) => <th key={h} className="px-4 py-3 font-semibold">{h}</th>)}</tr></thead>
        <tbody className="divide-y divide-slate-200">{rows.map(renderDesktop)}</tbody>
      </table>
    </div>
    <div className="mobile-record-list space-y-3 md:hidden">
      {rows.length === 0 ? <div className="rounded-2xl border border-slate-200 bg-white p-4 text-sm text-slate-500 shadow-sm">No records found.</div> : rows.map(renderMobile)}
    </div>
  </>;
}

function MobileCard({ kicker, title, subtitle, status, details = [], actions, onClick, danger }) {
  return <div onClick={onClick} className={`mobile-record-card rounded-2xl border border-slate-200 bg-white p-3 shadow-sm ${onClick ? "cursor-pointer active:bg-slate-50" : ""} ${danger ? "border-red-100" : ""}`}>
    <div className="mobile-record-header">
      <div className="min-w-0 flex-1">
        <p className="mobile-record-kicker">{kicker}</p>
        <div className="mobile-record-title">{title}</div>
        {subtitle && <div className="mobile-record-subtitle">{subtitle}</div>}
      </div>
      {status && <div className="mobile-record-status">{status}</div>}
    </div>
    {details.length > 0 && <div className="mobile-record-details">{details.map(([label, value]) => <div key={label} className="mobile-detail-tile"><p className="mobile-cell-label">{label}</p><div className="mobile-cell-value">{value}</div></div>)}</div>}
    {actions && <div className="mobile-record-actions">{actions}</div>}
  </div>;
}

function Modal({ title, subtitle, children, onClose, max = "max-w-4xl" }) {
  return <div className="app-modal-overlay fixed inset-0 z-50 flex items-center justify-center p-4" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
    <div className={`app-modal-shell max-h-[90vh] w-full ${max} overflow-y-auto rounded-3xl`}>
      <div className="app-modal-header sticky top-0 z-10 flex items-start justify-between gap-4 px-6 py-5">
        <div><h3 className="text-2xl font-bold">{title}</h3>{subtitle && <p className="mt-1 text-sm text-slate-500">{subtitle}</p>}</div>
        <button onClick={onClose} aria-label="Close popup" className="app-modal-close flex h-9 w-9 shrink-0 items-center justify-center rounded-full text-sm font-bold">X</button>
      </div>
      <div className="app-modal-body px-6 pb-6">{children}</div>
    </div>
  </div>;
}

function ConfirmDialog({ dialog, onCancel, onConfirm }) {
  const isDanger = dialog.tone !== "normal";
  return <div className="app-modal-overlay fixed inset-0 z-[60] flex items-center justify-center p-4" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}>
    <div className="app-confirm-shell w-full max-w-lg rounded-3xl">
      <div className="p-6">
        <div className={`mb-4 flex h-12 w-12 items-center justify-center rounded-2xl ${isDanger ? "bg-red-100 text-red-700" : "bg-blue-100 text-blue-700"}`}><span className="text-2xl font-black">!</span></div>
        <h3 className="text-2xl font-bold text-slate-900">{dialog.title}</h3>
        <p className="mt-3 whitespace-pre-line text-sm leading-relaxed text-slate-600">{dialog.message}</p>
      </div>
      <div className="app-modal-actions flex flex-wrap justify-end gap-2 px-6 py-4">
        <button type="button" onClick={onCancel} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button>
        <button type="button" onClick={onConfirm} className={`rounded-2xl px-5 py-2.5 text-sm font-semibold text-white ${isDanger ? "bg-red-700 hover:bg-red-800" : "bg-slate-900 hover:bg-slate-800"}`}>{dialog.confirmText}</button>
      </div>
    </div>
  </div>;
}

export default function ITAssetDashboardPreview() {
  const [users, setUsers] = useState(() => load(LS.users, defaultUsers));
  const [profiles, setProfiles] = useState(() => load(LS.profiles, defaultUserProfiles));
  const [groups, setGroups] = useState(() => load(LS.groups, defaultGroups));
  const [locations, setLocations] = useState(() => load(LS.locations, defaultLocations));
  const [branches, setBranches] = useState(() => load(LS.branches, defaultBranches));
  const [assets, setAssets] = useState(() => load(LS.assets, defaultAssets));
  const [networks, setNetworks] = useState(() => load(LS.networks, defaultNetworkSetups));
  const [handovers, setHandovers] = useState(() => load(LS.handovers, defaultHandovers));
  const [pullouts, setPullouts] = useState(() => load(LS.pullouts, defaultPullouts));
  const [faults, setFaults] = useState(() => load(LS.faults, defaultFaultLogs));
  const [assetProducts, setAssetProducts] = useState(() => load(LS.assetProducts, defaultAssetProducts).map(normalizeAssetProduct));
  const [assetCategories, setAssetCategories] = useState(() => mergeAssetCategories(load(LS.assetCategories, defaultAssetCategories), load(LS.assetProducts, defaultAssetProducts)));
  const [assetTemplates, setAssetTemplates] = useState(() => load(LS.assetTemplates, defaultAssetTemplates));
  const [credentials, setCredentials] = useState(() => load(LS.credentials, []).map(normalizeCredentialRecord));
  const [vaultSettings, setVaultSettings] = useState(() => load(LS.vaultSettings, null));
  const [resetRequests, setResetRequests] = useState(() => load(LS.resetRequests, []));
  const [vaultUnlocked, setVaultUnlocked] = useState(false);
  const vaultKeyRef = useRef(null);
  const [theme, setTheme] = useState(() => localStorage.getItem(LS.theme) || "light");

  const [currentUser, setCurrentUser] = useState(null);
  const [loginForm, setLoginForm] = useState({ username: "", password: "" });
  const [loginError, setLoginError] = useState("");
  const [loginStatus, setLoginStatus] = useState("idle");
  const [forgotOpen, setForgotOpen] = useState(false);
  const [forgotForm, setForgotForm] = useState({ username: "", contact: "", message: "" });
  const [forgotError, setForgotError] = useState("");
  const [forgotSuccess, setForgotSuccess] = useState("");
  const [forgotSubmitting, setForgotSubmitting] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [branchSubTab, setBranchSubTab] = useState("overview");
  const [selectedBranch, setSelectedBranch] = useState(() => branches[0] || defaultBranches[0]);
  const [selectedContributor, setSelectedContributor] = useState("Demo Technician");
  const [search, setSearch] = useState({ branch: "", asset: "", branchAsset: "", cctv: "", fault: "", location: "" });
  const [faultView, setFaultView] = useState("overall");
  const [reportRange, setReportRange] = useState({ start: "", end: "" });
  const [modal, setModal] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [formError, setFormError] = useState("");
  const [editId, setEditId] = useState(null);
  const [forms, setForms] = useState({ branch: emptyBranch, location: emptyLocation, asset: emptyAsset, product: emptyAssetProduct, assetTemplate: emptyAssetTemplate, network: emptyNetwork, handover: emptyHandover, pullout: emptyPullout, fault: emptyFault, profile: emptyProfile, groupBranding: emptyGroupBranding, credential: emptyCredential, cctv: { nvrUsername: "admin", nvrPassword: "", recordingStartDate: "", recordingEndDate: today(), recordingDays: "", lastRecordingReviewDate: today(), reviewedBy: "Demo Technician", reviewStatus: "Review Pending", issueFound: "", actionRequired: "", nextReviewDate: "", remarks: "" }, newUser: { displayName: "", username: "", password: "", role: "user" }, adminPassword: { currentPassword: "", newPassword: "", confirmPassword: "" } });
  const restoreInputRef = useRef(null);
  const excelWorkbookRef = useRef(null);
  const branchWorkbookRef = useRef(null);
  const locationWorkbookRef = useRef(null);
  const productWorkbookRef = useRef(null);
  const credentialImportRef = useRef(null);
  const [excelImport, setExcelImport] = useState(emptyExcelImport);
  const [branchImport, setBranchImport] = useState(emptyBranchImport);
  const [locationImport, setLocationImport] = useState(emptyLocationImport);
  const [productImport, setProductImport] = useState(emptyProductImport);
  const [templateBuilder, setTemplateBuilder] = useState({ type: "asset", selected: defaultTemplateSelected("asset") });
  const [serverLoaded, setServerLoaded] = useState(false);
  const [serverMessage, setServerMessage] = useState("Loading shared database...");

  useEffect(() => save(LS.users, users), [users]);
  useEffect(() => save(LS.profiles, profiles), [profiles]);
  useEffect(() => save(LS.groups, groups), [groups]);
  useEffect(() => save(LS.locations, locations), [locations]);
  useEffect(() => save(LS.branches, branches), [branches]);
  useEffect(() => save(LS.assets, assets), [assets]);
  useEffect(() => save(LS.networks, networks), [networks]);
  useEffect(() => save(LS.handovers, handovers), [handovers]);
  useEffect(() => save(LS.pullouts, pullouts), [pullouts]);
  useEffect(() => save(LS.faults, faults), [faults]);
  useEffect(() => save(LS.assetProducts, assetProducts), [assetProducts]);
  useEffect(() => save(LS.assetCategories, assetCategories), [assetCategories]);
  useEffect(() => save(LS.assetTemplates, assetTemplates), [assetTemplates]);
  useEffect(() => save(LS.credentials, credentials), [credentials]);
  useEffect(() => save(LS.vaultSettings, vaultSettings), [vaultSettings]);
  useEffect(() => save(LS.resetRequests, resetRequests), [resetRequests]);
  useEffect(() => localStorage.setItem(LS.theme, theme), [theme]);

  function buildDatabaseSnapshot() {
    return {
      appName: "Branch IT Control Dashboard",
      version: "v8.47-modal-contrast-upgrade",
      exportedAt: new Date().toISOString(),
      users,
      userProfiles: profiles,
      groups,
      locations,
      branches,
      assets,
      networkSetups: networks,
      handovers,
      pullouts,
      faultLogs: faults,
      assetProducts,
      assetCategories,
      assetTemplates,
      credentials,
      vaultSettings,
      resetRequests,
    };
  }

  function applyDatabaseSnapshot(data, sourceLabel = "shared database") {
    if (!data || !Array.isArray(data.branches)) return false;
    const importedBranches = data.branches.length > 0 ? data.branches : defaultBranches;
    setUsers(Array.isArray(data.users) && data.users.length > 0 ? data.users : defaultUsers);
    setProfiles(Array.isArray(data.userProfiles) && data.userProfiles.length > 0 ? data.userProfiles : defaultUserProfiles);
    setGroups(Array.isArray(data.groups) && data.groups.length > 0 ? data.groups : defaultGroups);
    setLocations(Array.isArray(data.locations) ? data.locations : defaultLocations);
    setBranches(importedBranches);
    setAssets(Array.isArray(data.assets) ? data.assets : []);
    setNetworks(Array.isArray(data.networkSetups) ? data.networkSetups : []);
    setHandovers(Array.isArray(data.handovers) ? data.handovers : []);
    setPullouts(Array.isArray(data.pullouts) ? data.pullouts : []);
    setFaults(Array.isArray(data.faultLogs) ? data.faultLogs : []);
    const loadedProducts = (Array.isArray(data.assetProducts) && data.assetProducts.length ? data.assetProducts : defaultAssetProducts).map(normalizeAssetProduct);
    setAssetProducts(loadedProducts);
    setAssetCategories(mergeAssetCategories(Array.isArray(data.assetCategories) ? data.assetCategories : defaultAssetCategories, loadedProducts));
    setAssetTemplates(Array.isArray(data.assetTemplates) && data.assetTemplates.length ? data.assetTemplates : defaultAssetTemplates);
    if (Array.isArray(data.credentials)) setCredentials(data.credentials.map(normalizeCredentialRecord));
    if (Array.isArray(data.resetRequests)) setResetRequests(data.resetRequests);
    if (data.vaultSettings && typeof data.vaultSettings === "object") {
      setVaultSettings(data.vaultSettings);
      vaultKeyRef.current = null;
      setVaultUnlocked(false);
    }
    setSelectedBranch(importedBranches[0] || defaultBranches[0]);
    setServerMessage(`Loaded ${importedBranches.length} branch record(s) from ${sourceLabel}.`);
    return true;
  }

async function syncFromServer(showAlert = false) {
  try {
    await ensureFirebaseLogin();

    const dbRef = doc(db, "appData", "main");
    const dbSnap = await getDoc(dbRef);

    if (dbSnap.exists()) {
      const data = dbSnap.data();
      const loaded = applyDatabaseSnapshot(data, "Firebase Firestore");
      setServerLoaded(true);

      if (showAlert && loaded) {
        alert(`Shared database synced from Firebase. Branches: ${data.branches?.length || 0}`);
      }
      return;
    }

    const snapshot = buildDatabaseSnapshot();

    await setDoc(dbRef, {
      ...snapshot,
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    });

    setServerLoaded(true);
    setServerMessage(`Created Firebase shared database. Branches: ${snapshot.branches.length}`);

    if (showAlert) {
      alert(`Firebase shared database created. Branches: ${snapshot.branches.length}`);
    }
  } catch (error) {
    setServerLoaded(false);
    setServerMessage(`Firebase sync failed: ${error.message}. Local browser data is still available.`);
    if (showAlert) alert(`Firebase sync failed.\n\n${error.message}`);
  }
}

async function saveToServer(showAlert = false) {
  try {
    await ensureFirebaseLogin();

    const snapshot = buildDatabaseSnapshot();
    const dbRef = doc(db, "appData", "main");

    await setDoc(dbRef, {
      ...snapshot,
      updatedAt: serverTimestamp(),
    }, { merge: true });

    setServerMessage(`Saved to Firebase shared database. Branches: ${snapshot.branches.length}`);

    if (showAlert) {
      alert("Database saved to Firebase successfully.");
    }
  } catch (error) {
    setServerMessage(`Save to Firebase failed: ${error.message}`);
    if (showAlert) alert(`Save to Firebase failed.\n\n${error.message}`);
  }
}

  useEffect(() => {
    syncFromServer(false);
  }, []);

  useEffect(() => {
    if (!serverLoaded) return;
    const handle = setTimeout(() => saveToServer(false), 700);
    return () => clearTimeout(handle);
  }, [serverLoaded, users, profiles, groups, locations, branches, assets, networks, handovers, pullouts, faults, assetProducts, assetCategories, assetTemplates, credentials, vaultSettings, resetRequests]);

  const isAdmin = currentUser?.role === "admin";
  const menuItems = [
    ["dashboard", "Dashboard", "dashboard"], ["contribution", "User Portfolio", "user"], ["groups", "Groups & Brands", "groups"], ["locations", "Locations", "location"], ["branches", "Branches", "branches"], ["assets", "IT Assets", "assets"], ["cctv", "CCTV / NVR", "cctv"], ["vault", "Credential Vault", "vault"], ["faults", "Fault Logs", "faults"], ["handover", "Handover Docs", "handover"], ["settings", "Settings", "settings"],
  ];
  const activePageTitle = tab === "branch-detail" ? selectedBranch?.branchName || "Branch Details" : (menuItems.find(([key]) => key === tab)?.[1] || "Dashboard");
  const pageDescriptions = {
    dashboard: "Overview of branch projects, locations and IT asset activity.", contribution: "Team portfolio, responsibilities and recorded work evidence.",
    groups: "Manage company groups, concepts and brand identities.", locations: "Maintain mall, site and operational location master data.",
    branches: "Browse and manage all branch IT project records.", assets: "Manage the complete IT asset register and standard templates.",
    cctv: "Review NVR recording health and camera deployment by operation type.", vault: "Encrypted device, cloud account and user-access credentials for the IT team.", faults: "Track branch IT faults, actions and closure status.",
    handover: "Access branch handover, pullout and global reports.", settings: "Manage users, appearance, backup and system controls.",
  };
  const selectedBranchAssets = useMemo(() => assets.filter((a) => a.branchId === selectedBranch?.id), [assets, selectedBranch]);
  const filteredSelectedBranchAssets = useMemo(() => {
    const k = search.branchAsset.toLowerCase();
    if (!k) return selectedBranchAssets;
    return selectedBranchAssets.filter((a) => `${a.category} ${a.deviceName} ${a.description} ${a.serialNumber} ${a.model} ${a.brand} ${a.ipAddress} ${a.macAddress} ${a.installedLocation} ${a.switchPort} ${a.status} ${a.remarks}`.toLowerCase().includes(k));
  }, [selectedBranchAssets, search.branchAsset]);
  const selectedNetwork = networks.find((n) => n.branchId === selectedBranch?.id);
  const selectedHandover = handovers.find((h) => h.branchId === selectedBranch?.id);
  const selectedPullout = pullouts.find((p) => p.branchId === selectedBranch?.id);
  const selectedProfile = profiles.find((p) => p.name === selectedContributor) || profiles[0] || defaultUserProfiles[0];
  const selectedGroupBranding = groups.find((g) => g.name === selectedBranch?.groupName) || defaultGroups[0];

  const totalCameras = assets.filter((a) => String(a.category || "").toLowerCase().includes("camera")).reduce((sum, a) => sum + Math.max(1, Number(a.quantity || 1)), 0);
  const totalNvr = assets.filter((a) => a.category === "NVR").length;
  const totalNvrBelow60 = assets.filter((a) => a.category === "NVR" && Number(a.recordingDays || 0) < 60).length;
  const totalPos = assets.filter((a) => a.category === "POS Terminal").length;
  const openFaultCount = faults.filter((f) => f.status !== "Closed").length;
  const pendingHandover = branches.filter((b) => Number(b.handover || 0) < 100).length;
  const completedProjects = branches.filter((b) => b.status === "Live" || Number(b.handover || 0) === 100);
  const ongoingProjects = branches.filter((b) => ["Installation Started", "Pending Handover", "Opening Today"].includes(b.status));
  const upcomingProjects = branches.filter((b) => b.status === "Planning");

  function requireAdmin() { if (isAdmin) return true; alert("View-only users cannot add, edit or delete records."); return false; }
  function askConfirm({ title, message, confirmText = "Confirm", tone = "danger" }, onConfirm) {
    setConfirmDialog({ title, message, confirmText, tone, onConfirm });
  }
  function updateForm(name, field, value) { setForms((f) => ({ ...f, [name]: { ...f[name], [field]: value } })); }
  function closeModal() { setModal(null); setFormError(""); setEditId(null); }
  useEffect(() => {
    if (!modal) return;
    const onKeyDown = (event) => { if (event.key === "Escape") closeModal(); };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [modal]);
  function openBranch(branch) { setSelectedBranch(branch); setBranchSubTab("overview"); setTab("branch-detail"); }
  async function login(e) {
    e.preventDefault();
    if (loginStatus === "loading") return;
    setLoginError("");
    setLoginStatus("loading");
    await new Promise((resolve) => setTimeout(resolve, 850));
    const u = users.find((x) => x.username.toLowerCase() === loginForm.username.trim().toLowerCase() && x.password === loginForm.password);
    if (!u) {
      setLoginError("Invalid username or password. Please check your credentials and try again.");
      setLoginStatus("idle");
      return;
    }
    setLoginStatus("success");
    await new Promise((resolve) => setTimeout(resolve, 450));
    setCurrentUser(u);
    const p = profiles.find((x) => x.username === u.username || x.name.toLowerCase() === String(u.displayName).toLowerCase());
    if (p) setSelectedContributor(p.name);
    setLoginForm({ username: "", password: "" });
    setLoginError("");
    setLoginStatus("idle");
  }
  function openForgotPassword() {
    setForgotForm((current) => ({ ...current, username: current.username || loginForm.username }));
    setForgotError("");
    setForgotSuccess("");
    setForgotOpen(true);
  }
  function closeForgotPassword() {
    if (forgotSubmitting) return;
    setForgotOpen(false);
    setForgotError("");
    setForgotSuccess("");
  }
  async function submitForgotPassword(e) {
    e.preventDefault();
    const username = forgotForm.username.trim().toLowerCase();
    const contact = forgotForm.contact.trim();
    if (!username || !contact) {
      setForgotError("Enter your username and registered email address or mobile number.");
      return;
    }
    setForgotError("");
    setForgotSubmitting(true);
    await new Promise((resolve) => setTimeout(resolve, 650));
    const account = users.find((user) => user.username.toLowerCase() === username);
    const profile = profiles.find((item) => item.username?.toLowerCase() === username);
    const request = {
      id: makeId(),
      username,
      displayName: account?.displayName || profile?.name || "Unverified account",
      contact,
      message: forgotForm.message.trim(),
      status: account ? "Pending" : "Needs Verification",
      requestedAt: new Date().toISOString(),
      resolvedAt: "",
      resolvedBy: "",
    };
    setResetRequests((items) => [request, ...items]);
    setForgotSubmitting(false);
    setForgotSuccess("Your password-reset request has been submitted. Please contact the system administrator for confirmation.");
    setForgotForm({ username: "", contact: "", message: "" });
  }
  function logout() { vaultKeyRef.current = null; setVaultUnlocked(false); setCurrentUser(null); setTab("dashboard"); setBranchSubTab("overview"); }

  function updateBranchLogo(file) {
    if (!requireAdmin()) return;
    if (!file) return;
    if (!file.type || !file.type.startsWith("image/")) {
      alert("Please upload a PNG/JPG image file.");
      return;
    }
    if (file.size > 2 * 1024 * 1024) {
      alert("Logo file is too large. Please upload an image below 2 MB.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const logoDataUrl = String(reader.result || "");
      const updated = { ...selectedBranch, logoDataUrl };
      setSelectedBranch(updated);
      setBranches((items) => items.map((branch) => branch.id === updated.id ? updated : branch));
    };
    reader.readAsDataURL(file);
  }

  function removeBranchLogo() {
    if (!requireAdmin()) return;
    const updated = { ...selectedBranch, logoDataUrl: "" };
    setSelectedBranch(updated);
    setBranches((items) => items.map((branch) => branch.id === updated.id ? updated : branch));
  }

  function exportDatabase() {
    const data = buildDatabaseSnapshot();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `branch-it-dashboard-backup-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.json`;
    a.click(); URL.revokeObjectURL(a.href);
  }
  function restoreDatabaseFile(e) {
    const file = e.target.files?.[0]; if (!file) return;
    const r = new FileReader();
    r.onload = async () => {
      try {
        const data = JSON.parse(String(r.result || "{}"));
        if (!Array.isArray(data.branches) || data.branches.length === 0) throw new Error("Invalid backup");
        const ok = confirm(`Restore database backup?\n\nFile: ${file.name}\nBranches: ${data.branches.length}\nAssets: ${(data.assets || []).length}\nFault Logs: ${(data.faultLogs || []).length}\n\nThis will replace current app data and save to shared data/db.json for mobile/Tailscale.`);
        if (!ok) return;
        applyDatabaseSnapshot(data, file.name);
        setServerLoaded(true);
        setTimeout(() => saveToServer(true), 150);
      } catch {
        alert("Restore failed. Please select a valid backup JSON file.");
      } finally {
        e.target.value = "";
      }
    };
    r.readAsText(file);
  }
  function resetData() { if (!requireAdmin()) return; if (confirm("Download backup before reset?")) exportDatabase(); const text = prompt("Danger Zone: type exactly RESET DEMO to reset data."); if (text !== "RESET DEMO") return alert("Reset cancelled."); if (!confirm("Final confirmation: reset all current data?")) return; setBranches(defaultBranches); setLocations(defaultLocations); setAssets(defaultAssets); setNetworks(defaultNetworkSetups); setHandovers(defaultHandovers); setPullouts(defaultPullouts); setFaults(defaultFaultLogs); setResetRequests([]); setSelectedBranch(defaultBranches[0]); setServerLoaded(true); setTimeout(() => saveToServer(true), 150); }

  const filteredBranches = useMemo(() => {
    const k = search.branch.toLowerCase();
    return [...branches].sort((a, b) => String(b.completedDate).localeCompare(String(a.completedDate))).filter((b) => `${b.groupName} ${b.conceptName} ${b.branchName} ${b.locationName} ${b.city} ${b.status}`.toLowerCase().includes(k));
  }, [branches, search.branch]);
  const enrichedAssets = useMemo(() => assets.map((a) => ({ ...a, branch: branches.find((b) => b.id === a.branchId) || {} })), [assets, branches]);
  const filteredAssets = useMemo(() => enrichedAssets.filter((a) => `${a.category} ${a.deviceName} ${a.serialNumber} ${a.ipAddress} ${a.branch.branchName} ${a.branch.locationName} ${a.brand}`.toLowerCase().includes(search.asset.toLowerCase())), [enrichedAssets, search.asset]);
  const nvrRecords = useMemo(() => enrichedAssets.filter((a) => a.category === "NVR" && `${a.deviceName} ${a.serialNumber} ${a.ipAddress} ${a.branch.branchName} ${a.branch.locationName}`.toLowerCase().includes(search.cctv.toLowerCase())), [enrichedAssets, search.cctv]);
  const enrichedFaults = useMemo(() => faults.map((f) => ({ ...f, branch: branches.find((b) => b.id === f.branchId) || {} })).sort((a, b) => `${b.faultDate} ${b.faultTime}`.localeCompare(`${a.faultDate} ${a.faultTime}`)), [faults, branches]);
  const filteredFaults = useMemo(() => enrichedFaults.filter((f) => `${f.branch.branchName} ${f.branch.locationName} ${f.category} ${f.title} ${f.attendedBy} ${f.actionTaken} ${f.status}`.toLowerCase().includes(search.fault.toLowerCase())), [enrichedFaults, search.fault]);
  const selectedBranchFaults = useMemo(() => faults.filter((f) => f.branchId === selectedBranch?.id).sort((a, b) => `${b.faultDate} ${b.faultTime}`.localeCompare(`${a.faultDate} ${a.faultTime}`)), [faults, selectedBranch]);

  const contributorProjects = branches.filter((b) => `${b.preparedBy || ""} ${b.notes || ""}`.toLowerCase().includes(selectedContributor.toLowerCase()));
  const contributorFaults = faults.filter((f) => `${f.attendedBy || ""} ${f.actionTaken || ""} ${f.remarks || ""}`.toLowerCase().includes(selectedContributor.toLowerCase()));
  const contributorHandovers = handovers.filter((h) => `${h.completedPersons || ""} ${h.inchargeName || ""}`.toLowerCase().includes(selectedContributor.toLowerCase()));
  const contributorPullouts = pullouts.filter((p) => `${p.removedBy || ""} ${p.inchargeName || ""}`.toLowerCase().includes(selectedContributor.toLowerCase()));
  const contributorCctvReviews = assets.filter((a) => a.category === "NVR" && `${a.reviewedBy || ""} ${a.remarks || ""}`.toLowerCase().includes(selectedContributor.toLowerCase()));

  useEffect(() => {
    setSearch({ branch: "", asset: "", branchAsset: "", cctv: "", fault: "", location: "" });
    setFaultView("overall");
    setReportRange({ start: "", end: "" });
  }, [tab]);
  useEffect(() => {
    setSearch((current) => ({ ...current, branchAsset: "" }));
  }, [branchSubTab]);

  function resetExcelImport() {
    excelWorkbookRef.current = null;
    setExcelImport(emptyExcelImport);
  }

  function loadExcelSheet(sheetName, workbook = excelWorkbookRef.current) {
    if (!workbook || !sheetName || !workbook.Sheets?.[sheetName]) return;
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false, blankrows: false });
    const headers = Array.from(new Set(rows.flatMap((row) => Object.keys(row).map((key) => String(key).trim()).filter(Boolean))));
    const mapping = detectExcelMapping(headers);
    setExcelImport((current) => ({ ...current, selectedSheet: sheetName, headers, rows, mapping, preview: null }));
  }

  async function handleExcelAssetFile(file) {
    if (!requireAdmin()) return;
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls") && !lower.endsWith(".csv")) {
      alert("Please upload an Excel or CSV file: .xlsx, .xls or .csv");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) throw new Error("No sheets found in Excel file.");
      excelWorkbookRef.current = workbook;
      const selectedSheet = sheetNames[0];
      setExcelImport((current) => ({ ...emptyExcelImport, duplicateMode: current.duplicateMode || "skip", fileName: file.name, sheetNames, selectedSheet }));
      loadExcelSheet(selectedSheet, workbook);
    } catch (error) {
      alert(`Excel import failed.\n\n${error.message || error}`);
    }
  }

  function updateExcelMapping(field, value) {
    setExcelImport((current) => ({ ...current, mapping: { ...current.mapping, [field]: value }, preview: null }));
  }

  function updateExcelDuplicateMode(value) {
    setExcelImport((current) => ({ ...current, duplicateMode: value, preview: null }));
  }

  function buildExcelImportPreview(showAlert = true) {
    if (!excelImport.rows.length) {
      if (showAlert) alert("Please upload an Excel file and select a sheet first.");
      return null;
    }
    const requiredMissing = excelImportFields
      .filter(([, , required]) => required)
      .filter(([field]) => !excelImport.mapping[field]);
    if (requiredMissing.length) {
      if (showAlert) alert(`Please map required columns first:\n\n${requiredMissing.map(([, label]) => label).join("\n")}`);
      return null;
    }

    const existingSerials = new Set(assets.map((asset) => normalizeSerial(asset.serialNumber)).filter(Boolean));
    const parsedAssets = [];
    let skippedEmpty = 0;
    let missingSerial = 0;
    let duplicates = 0;

    excelImport.rows.forEach((row, index) => {
      const serialRaw = excelValue(row, excelImport.mapping.serialNumber);
      const categoryRaw = excelValue(row, excelImport.mapping.category);
      const deviceNameRaw = excelValue(row, excelImport.mapping.deviceName);
      const description = excelValue(row, excelImport.mapping.description);
      const brand = excelValue(row, excelImport.mapping.brand);
      const model = excelValue(row, excelImport.mapping.model);
      const quantity = cleanExcelQuantity(excelValue(row, excelImport.mapping.quantity));
      const ipAddress = excelValue(row, excelImport.mapping.ipAddress);
      const macAddress = excelValue(row, excelImport.mapping.macAddress);
      const installedLocation = excelValue(row, excelImport.mapping.installedLocation);
      const switchPort = excelValue(row, excelImport.mapping.switchPort);
      const channelNumber = excelValue(row, excelImport.mapping.channelNumber);
      const cameraView = excelValue(row, excelImport.mapping.cameraView);
      const warrantyStart = cleanDateText(excelValue(row, excelImport.mapping.warrantyStart));
      const warrantyEnd = cleanDateText(excelValue(row, excelImport.mapping.warrantyEnd));
      const status = excelValue(row, excelImport.mapping.status);
      const remarks = cleanImportedPlaceholder(excelValue(row, excelImport.mapping.remarks));
      const nvrUsername = excelValue(row, excelImport.mapping.nvrUsername);
      const nvrPassword = excelValue(row, excelImport.mapping.nvrPassword);
      const recordingDays = excelValue(row, excelImport.mapping.recordingDays);
      const recordingStartDate = cleanDateText(excelValue(row, excelImport.mapping.recordingStartDate));
      const recordingEndDate = cleanDateText(excelValue(row, excelImport.mapping.recordingEndDate));
      const lastRecordingReviewDate = cleanDateText(excelValue(row, excelImport.mapping.lastRecordingReviewDate));
      const reviewedBy = excelValue(row, excelImport.mapping.reviewedBy);
      const reviewStatus = excelValue(row, excelImport.mapping.reviewStatus);
      const issueFound = excelValue(row, excelImport.mapping.issueFound);
      const actionRequired = excelValue(row, excelImport.mapping.actionRequired);
      const nextReviewDate = cleanDateText(excelValue(row, excelImport.mapping.nextReviewDate));

      if (!serialRaw && !categoryRaw && !deviceNameRaw && !description && !brand && !model && !installedLocation) {
        skippedEmpty += 1;
        return;
      }

      const generatedSerial = serialRaw || `NO-SERIAL-${selectedBranch.id}-${index + 1}`;
      if (!serialRaw) missingSerial += 1;
      const duplicate = serialRaw ? existingSerials.has(normalizeSerial(serialRaw)) : false;
      if (duplicate) duplicates += 1;
      const calculatedDays = countRecordingDays(recordingStartDate, recordingEndDate);

      parsedAssets.push({
        sourceRow: index + 2,
        duplicate,
        category: mapExcelCategory(categoryRaw),
        customCategory: "",
        deviceName: deviceNameRaw || description || categoryRaw || "Imported Asset",
        description: description || "",
        quantity: quantity || "1",
        brand: brand || "",
        model: model || "",
        serialNumber: generatedSerial,
        ipAddress: ipAddress || "",
        macAddress: macAddress || "",
        installedLocation: installedLocation || "",
        switchPort: switchPort || "",
        channelNumber: channelNumber || "",
        cameraView: cameraView || "",
        warrantyStart: warrantyStart || "",
        warrantyEnd: warrantyEnd || "",
        status: status || "Installed",
        remarks: remarks || "",
        nvrUsername: nvrUsername || "",
        nvrPassword: nvrPassword || "",
        recordingStartDate: recordingStartDate || "",
        recordingEndDate: recordingEndDate || "",
        recordingDays: calculatedDays || recordingDays || "",
        lastRecordingReviewDate: lastRecordingReviewDate || "",
        reviewedBy: reviewedBy || "",
        reviewStatus: reviewStatus || (calculatedDays ? getAutoNvrStatus(calculatedDays) : ""),
        issueFound: issueFound || "",
        actionRequired: actionRequired || "",
        nextReviewDate: nextReviewDate || "",
      });
    });

    const importable = parsedAssets.filter((asset) => excelImport.duplicateMode !== "skip" || !asset.duplicate);
    const preview = {
      totalRows: excelImport.rows.length,
      skippedEmpty,
      missingSerial,
      duplicates,
      parsedAssets,
      importable,
      newCount: parsedAssets.filter((asset) => !asset.duplicate).length,
      updateCount: excelImport.duplicateMode === "update" ? parsedAssets.filter((asset) => asset.duplicate).length : 0,
      skipCount: excelImport.duplicateMode === "skip" ? parsedAssets.filter((asset) => asset.duplicate).length : 0,
    };
    setExcelImport((current) => ({ ...current, preview }));
    return preview;
  }

  function confirmExcelAssetImport() {
    if (!requireAdmin()) return;
    const preview = excelImport.preview || buildExcelImportPreview(false);
    if (!preview || !preview.importable.length) {
      alert("No assets available to import. Check the selected sheet, column mapping and duplicate option.");
      return;
    }

    const ok = confirm(`Import assets into branch?\n\nBranch: ${selectedBranch.branchName} - ${selectedBranch.locationName}\nSheet: ${excelImport.selectedSheet}\nAssets to save: ${preview.importable.length}\nDuplicates found: ${preview.duplicates}\nDuplicate mode: ${excelImport.duplicateMode}\n\nContinue?`);
    if (!ok) return;

    setAssets((current) => {
      let next = [...current];
      const serialIndex = new Map(next.map((asset, index) => [normalizeSerial(asset.serialNumber), index]).filter(([serial]) => serial));
      preview.parsedAssets.forEach((asset, index) => {
        const serial = normalizeSerial(asset.serialNumber);
        const existingIndex = serial ? serialIndex.get(serial) : undefined;
        if (existingIndex !== undefined && excelImport.duplicateMode === "skip") return;
        if (existingIndex !== undefined && excelImport.duplicateMode === "update") {
          next[existingIndex] = { ...next[existingIndex], ...asset, id: next[existingIndex].id, branchId: selectedBranch.id };
          return;
        }
        const savedAsset = { ...asset, id: makeId() + index, branchId: selectedBranch.id };
        next = [savedAsset, ...next];
      });
      return next;
    });

    alert(`Excel import completed.\n\nImported/Saved: ${preview.importable.length}\nSkipped duplicates: ${preview.skipCount}`);
    resetExcelImport();
    closeModal();
    setBranchSubTab("assets");
  }


  function resetProductImport() {
    productWorkbookRef.current = null;
    setProductImport(emptyProductImport);
  }

  function loadProductExcelSheet(sheetName, workbook = productWorkbookRef.current) {
    if (!workbook || !sheetName || !workbook.Sheets?.[sheetName]) return;
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false, blankrows: false });
    const headers = Array.from(new Set(rows.flatMap((row) => Object.keys(row).map((key) => String(key).trim()).filter(Boolean))));
    const mapping = detectProductMapping(headers);
    setProductImport((current) => ({ ...current, selectedSheet: sheetName, headers, rows, mapping, preview: null }));
  }

  async function handleProductExcelFile(file) {
    if (!requireAdmin()) return;
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls") && !lower.endsWith(".csv")) {
      alert("Please upload an Excel or CSV file: .xlsx, .xls or .csv");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) throw new Error("No sheets found in the uploaded file.");
      productWorkbookRef.current = workbook;
      const selectedSheet = sheetNames[0];
      setProductImport((current) => ({ ...emptyProductImport, duplicateMode: current.duplicateMode || "skip", fileName: file.name, sheetNames, selectedSheet }));
      loadProductExcelSheet(selectedSheet, workbook);
    } catch (error) {
      alert(`Product Excel import failed.

${error.message || error}`);
    }
  }

  function updateProductImportMapping(field, value) {
    setProductImport((current) => ({ ...current, mapping: { ...current.mapping, [field]: value }, preview: null }));
  }

  function updateProductImportDuplicateMode(value) {
    setProductImport((current) => ({ ...current, duplicateMode: value, preview: null }));
  }

  function buildProductImportPreview(showAlert = true) {
    if (!productImport.rows.length) {
      if (showAlert) alert("Please upload the product Excel template first.");
      return null;
    }
    const requiredMissing = productMasterFields
      .filter(([, , required]) => required)
      .filter(([field]) => !productImport.mapping[field]);
    if (requiredMissing.length) {
      if (showAlert) alert(`Please map required columns first:

${requiredMissing.map(([, label]) => label).join("\n")}`);
      return null;
    }

    const existingIndex = new Map(assetProducts.map((item, index) => [normalizeProductKey(item), index]));
    const parsedProducts = [];
    let skippedEmpty = 0;
    let invalid = 0;
    let duplicates = 0;

    productImport.rows.forEach((row, index) => {
      const productName = excelValue(row, productImport.mapping.productName);
      const categoryRaw = excelValue(row, productImport.mapping.category);
      const brand = excelValue(row, productImport.mapping.brand);
      const model = excelValue(row, productImport.mapping.model);
      const description = excelValue(row, productImport.mapping.description);
      const defaultQuantity = cleanExcelQuantity(excelValue(row, productImport.mapping.defaultQuantity));
      const defaultLocation = excelValue(row, productImport.mapping.defaultLocation);
      const defaultStatus = excelValue(row, productImport.mapping.defaultStatus) || "Installed";
      const remarks = cleanImportedPlaceholder(excelValue(row, productImport.mapping.remarks));

      if (!productName && !categoryRaw && !brand && !model && !description) {
        skippedEmpty += 1;
        return;
      }
      if (!productName || !categoryRaw) {
        invalid += 1;
        return;
      }

      const product = {
        sourceRow: index + 2,
        productName: productName.trim(),
        category: mapExcelCategory(categoryRaw),
        brand: brand || "",
        model: model || "",
        description: description || "",
        defaultQuantity: defaultQuantity || "1",
        defaultLocation: defaultLocation || "",
        defaultStatus,
        isActive: true,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        remarks: remarks || "",
      };
      const key = normalizeProductKey(product);
      const duplicate = existingIndex.has(key);
      if (duplicate) duplicates += 1;
      parsedProducts.push({ ...product, duplicate, existingIndex: duplicate ? existingIndex.get(key) : -1 });
    });

    const importable = parsedProducts.filter((item) => productImport.duplicateMode !== "skip" || !item.duplicate);
    const preview = {
      totalRows: productImport.rows.length,
      skippedEmpty,
      invalid,
      duplicates,
      parsedProducts,
      importable,
      newCount: parsedProducts.filter((item) => !item.duplicate).length,
      updateCount: productImport.duplicateMode === "update" ? parsedProducts.filter((item) => item.duplicate).length : 0,
      skipCount: productImport.duplicateMode === "skip" ? parsedProducts.filter((item) => item.duplicate).length : 0,
    };
    setProductImport((current) => ({ ...current, preview }));
    return preview;
  }

  function confirmProductImport() {
    if (!requireAdmin()) return;
    const preview = productImport.preview || buildProductImportPreview(false);
    if (!preview || !preview.importable.length) {
      alert("No products are available to import. Check the sheet, mapping and duplicate option.");
      return;
    }
    const ok = confirm(`Import products into Asset Product List?

Sheet: ${productImport.selectedSheet}
Products to save: ${preview.importable.length}
Duplicates found: ${preview.duplicates}
Invalid rows: ${preview.invalid}
Duplicate mode: ${productImport.duplicateMode}

Continue?`);
    if (!ok) return;

    const importedCategoryNames = Array.from(new Set(preview.parsedProducts.map((item) => item.category).filter(Boolean)));
    setAssetCategories((current) => {
      const existing = new Set(current.map((item) => normalizeExcelText(item.name)));
      const additions = importedCategoryNames.filter((name) => !existing.has(normalizeExcelText(name))).map((name) => ({ id: makeId() + name, name, description: "Added from Product Excel import", isActive: true }));
      return additions.length ? [...current, ...additions].sort((a, b) => a.name.localeCompare(b.name)) : current;
    });

    setAssetProducts((current) => {
      let next = [...current];
      let keyIndex = new Map(next.map((item, index) => [normalizeProductKey(item), index]));
      preview.parsedProducts.forEach((product, index) => {
        const key = normalizeProductKey(product);
        const existingAt = keyIndex.get(key);
        if (existingAt !== undefined && productImport.duplicateMode === "skip") return;
        if (existingAt !== undefined && productImport.duplicateMode === "update") {
          const cleanProduct = { ...product };
          delete cleanProduct.sourceRow;
          delete cleanProduct.duplicate;
          delete cleanProduct.existingIndex;
          next[existingAt] = normalizeAssetProduct({ ...next[existingAt], ...cleanProduct, id: next[existingAt].id, updatedAt: new Date().toISOString() });
          return;
        }
        const cleanProduct = { ...product, id: makeId() + index };
        delete cleanProduct.sourceRow;
        delete cleanProduct.duplicate;
        delete cleanProduct.existingIndex;
        next = [normalizeAssetProduct(cleanProduct), ...next];
        keyIndex = new Map(next.map((item, itemIndex) => [normalizeProductKey(item), itemIndex]));
      });
      return next;
    });

    alert(`Product import completed.

Saved: ${preview.importable.length}
Skipped duplicates: ${preview.skipCount}
Invalid rows: ${preview.invalid}`);
    resetProductImport();
    closeModal();
  }



  function resetLocationImport() {
    locationWorkbookRef.current = null;
    setLocationImport(emptyLocationImport);
  }

  function loadLocationExcelSheet(sheetName, workbook = locationWorkbookRef.current) {
    if (!workbook || !sheetName || !workbook.Sheets?.[sheetName]) return;
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false, blankrows: false });
    const headers = Array.from(new Set(rows.flatMap((row) => Object.keys(row).map((key) => String(key).trim()).filter(Boolean))));
    const mapping = detectLocationMapping(headers);
    setLocationImport((current) => ({ ...current, selectedSheet: sheetName, headers, rows, mapping, preview: null }));
  }

  async function handleLocationExcelFile(file) {
    if (!requireAdmin()) return;
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls") && !lower.endsWith(".csv")) {
      alert("Please upload an Excel or CSV file: .xlsx, .xls or .csv");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) throw new Error("No sheets found in Excel file.");
      locationWorkbookRef.current = workbook;
      const selectedSheet = sheetNames[0];
      setLocationImport((current) => ({ ...emptyLocationImport, duplicateMode: current.duplicateMode || "skip", fileName: file.name, sheetNames, selectedSheet }));
      loadLocationExcelSheet(selectedSheet, workbook);
    } catch (error) {
      alert(`Location Excel import failed.\n\n${error.message || error}`);
    }
  }

  function updateLocationImportMapping(field, value) {
    setLocationImport((current) => ({ ...current, mapping: { ...current.mapping, [field]: value }, preview: null }));
  }

  function updateLocationImportDuplicateMode(value) {
    setLocationImport((current) => ({ ...current, duplicateMode: value, preview: null }));
  }

  function buildLocationImportPreview(showAlert = true) {
    if (!locationImport.rows.length) {
      if (showAlert) alert("Please upload an Excel file and select a sheet first.");
      return null;
    }
    const requiredMissing = locationImportFields.filter(([, , required]) => required).filter(([field]) => !locationImport.mapping[field]);
    if (requiredMissing.length) {
      if (showAlert) alert(`Please map required columns first:\n\n${requiredMissing.map(([, label]) => label).join("\n")}`);
      return null;
    }

    const existingKeys = new Set(locations.map((location) => normalizeLocationKey(location.name, location.city, location.emirate)));
    const parsedLocations = [];
    let skippedEmpty = 0;
    let missingRequired = 0;
    let duplicates = 0;

    locationImport.rows.forEach((row, index) => {
      const name = excelValue(row, locationImport.mapping.name);
      const city = excelValue(row, locationImport.mapping.city);
      const emirate = excelValue(row, locationImport.mapping.emirate);
      const type = cleanLocationType(excelValue(row, locationImport.mapping.type));
      const notes = excelValue(row, locationImport.mapping.notes);

      if (!name && !city && !emirate && !type && !notes) {
        skippedEmpty += 1;
        return;
      }
      if (!name || !city || !emirate) {
        missingRequired += 1;
        return;
      }

      const duplicate = existingKeys.has(normalizeLocationKey(name, city, emirate));
      if (duplicate) duplicates += 1;
      parsedLocations.push({
        sourceRow: index + 2,
        duplicate,
        name,
        city,
        emirate,
        type,
        notes: notes || `Imported from Excel: ${locationImport.fileName} / ${locationImport.selectedSheet}`,
      });
    });

    const importable = parsedLocations.filter((location) => locationImport.duplicateMode !== "skip" || !location.duplicate);
    const preview = {
      totalRows: locationImport.rows.length,
      skippedEmpty,
      missingRequired,
      duplicates,
      parsedLocations,
      importable,
      newCount: parsedLocations.filter((location) => !location.duplicate).length,
      updateCount: locationImport.duplicateMode === "update" ? parsedLocations.filter((location) => location.duplicate).length : 0,
      skipCount: locationImport.duplicateMode === "skip" ? parsedLocations.filter((location) => location.duplicate).length : 0,
    };
    setLocationImport((current) => ({ ...current, preview }));
    return preview;
  }

  function confirmLocationImport() {
    if (!requireAdmin()) return;
    const preview = locationImport.preview || buildLocationImportPreview(false);
    if (!preview || !preview.importable.length) {
      alert("No locations available to import. Check sheet, mapping and duplicate option.");
      return;
    }
    const ok = confirm(`Import locations from Excel?\n\nSheet: ${locationImport.selectedSheet}\nLocations to save: ${preview.importable.length}\nDuplicates found: ${preview.duplicates}\nDuplicate mode: ${locationImport.duplicateMode}\n\nContinue?`);
    if (!ok) return;

    setLocations((current) => {
      let next = [...current];
      const keyIndex = new Map(next.map((location, index) => [normalizeLocationKey(location.name, location.city, location.emirate), index]));
      preview.parsedLocations.forEach((location, index) => {
        const key = normalizeLocationKey(location.name, location.city, location.emirate);
        const existingIndex = keyIndex.get(key);
        if (existingIndex !== undefined && locationImport.duplicateMode === "skip") return;
        if (existingIndex !== undefined && locationImport.duplicateMode === "update") {
          next[existingIndex] = { ...next[existingIndex], ...location, id: next[existingIndex].id };
          return;
        }
        next = [{ ...location, id: makeId() + index }, ...next];
      });
      return next;
    });

    alert(`Location import completed.\n\nSaved: ${preview.importable.length}\nSkipped duplicates: ${preview.skipCount}`);
    resetLocationImport();
    closeModal();
    setTab("locations");
  }

  function resetBranchImport() {
    branchWorkbookRef.current = null;
    setBranchImport(emptyBranchImport);
  }

  function loadBranchExcelSheet(sheetName, workbook = branchWorkbookRef.current) {
    if (!workbook || !sheetName || !workbook.Sheets?.[sheetName]) return;
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false, blankrows: false });
    const headers = Array.from(new Set(rows.flatMap((row) => Object.keys(row).map((key) => String(key).trim()).filter(Boolean))));
    const mapping = detectBranchMapping(headers);
    setBranchImport((current) => ({ ...current, selectedSheet: sheetName, headers, rows, mapping, preview: null }));
  }

  async function handleBranchExcelFile(file) {
    if (!requireAdmin()) return;
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls") && !lower.endsWith(".csv")) {
      alert("Please upload an Excel or CSV file: .xlsx, .xls or .csv");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) throw new Error("No sheets found in Excel file.");
      branchWorkbookRef.current = workbook;
      const selectedSheet = sheetNames[0];
      setBranchImport((current) => ({ ...emptyBranchImport, duplicateMode: current.duplicateMode || "skip", fileName: file.name, sheetNames, selectedSheet }));
      loadBranchExcelSheet(selectedSheet, workbook);
    } catch (error) {
      alert(`Branch Excel import failed.\n\n${error.message || error}`);
    }
  }

  function updateBranchImportMapping(field, value) {
    setBranchImport((current) => ({ ...current, mapping: { ...current.mapping, [field]: value }, preview: null }));
  }

  function updateBranchImportDuplicateMode(value) {
    setBranchImport((current) => ({ ...current, duplicateMode: value, preview: null }));
  }

  function buildBranchImportPreview(showAlert = true) {
    if (!branchImport.rows.length) {
      if (showAlert) alert("Please upload an Excel file and select a sheet first.");
      return null;
    }
    const requiredMissing = branchImportFields.filter(([, , required]) => required).filter(([field]) => !branchImport.mapping[field]);
    if (requiredMissing.length) {
      if (showAlert) alert(`Please map required columns first:\n\n${requiredMissing.map(([, label]) => label).join("\n")}`);
      return null;
    }

    const existingKeys = new Set(branches.map((branch) => normalizeBranchKey(branch.branchName, branch.locationName)));
    const parsedBranches = [];
    let skippedEmpty = 0;
    let missingRequired = 0;
    let duplicates = 0;

    branchImport.rows.forEach((row, index) => {
      const groupName = excelValue(row, branchImport.mapping.groupName);
      const conceptName = excelValue(row, branchImport.mapping.conceptName);
      const branchName = excelValue(row, branchImport.mapping.branchName);
      const locationName = excelValue(row, branchImport.mapping.locationName);
      const city = excelValue(row, branchImport.mapping.city);
      const emirate = excelValue(row, branchImport.mapping.emirate);
      const branchType = cleanBranchType(excelValue(row, branchImport.mapping.branchType));
      const openingDate = cleanDateText(excelValue(row, branchImport.mapping.openingDate));
      const completedDate = cleanDateText(excelValue(row, branchImport.mapping.completedDate));
      const status = cleanBranchStatus(excelValue(row, branchImport.mapping.status));
      const branchManager = excelValue(row, branchImport.mapping.branchManager);
      const contactNumber = excelValue(row, branchImport.mapping.contactNumber);
      const preparedBy = excelValue(row, branchImport.mapping.preparedBy);
      const notes = excelValue(row, branchImport.mapping.notes);

      if (!groupName && !conceptName && !branchName && !locationName && !branchType) {
        skippedEmpty += 1;
        return;
      }
      if (!groupName || !branchName || !locationName || !branchType) {
        missingRequired += 1;
        return;
      }

      const loc = locations.find((item) => normalizeExcelText(item.name) === normalizeExcelText(locationName));
      const duplicate = existingKeys.has(normalizeBranchKey(branchName, locationName));
      if (duplicate) duplicates += 1;

      parsedBranches.push({
        sourceRow: index + 2,
        duplicate,
        groupName,
        conceptName: conceptName || groupName,
        branchName,
        locationName,
        city: city || loc?.city || "To be updated",
        emirate: emirate || loc?.emirate || "To be updated",
        branchType,
        openingDate,
        completedDate,
        status,
        branchManager: branchManager || "To be updated",
        contactNumber: contactNumber || "To be updated",
        preparedBy: preparedBy || "Demo Technician",
        notes: notes || `Imported from Excel: ${branchImport.fileName} / ${branchImport.selectedSheet}`,
        handover: 0,
      });
    });

    const importable = parsedBranches.filter((branch) => branchImport.duplicateMode !== "skip" || !branch.duplicate);
    const preview = {
      totalRows: branchImport.rows.length,
      skippedEmpty,
      missingRequired,
      duplicates,
      parsedBranches,
      importable,
      newCount: parsedBranches.filter((branch) => !branch.duplicate).length,
      updateCount: branchImport.duplicateMode === "update" ? parsedBranches.filter((branch) => branch.duplicate).length : 0,
      skipCount: branchImport.duplicateMode === "skip" ? parsedBranches.filter((branch) => branch.duplicate).length : 0,
    };
    setBranchImport((current) => ({ ...current, preview }));
    return preview;
  }

  function confirmBranchImport() {
    if (!requireAdmin()) return;
    const preview = branchImport.preview || buildBranchImportPreview(false);
    if (!preview || !preview.importable.length) {
      alert("No branches available to import. Check sheet, mapping and duplicate option.");
      return;
    }
    const ok = confirm(`Import branches from Excel?\n\nSheet: ${branchImport.selectedSheet}\nBranches to save: ${preview.importable.length}\nDuplicates found: ${preview.duplicates}\nDuplicate mode: ${branchImport.duplicateMode}\n\nContinue?`);
    if (!ok) return;

    const importableLocationNames = new Set(preview.importable.map((branch) => normalizeExcelText(branch.locationName)).filter(Boolean));
    setLocations((current) => {
      const existing = new Set(current.map((loc) => normalizeExcelText(loc.name)));
      const additions = [];
      preview.importable.forEach((branch, index) => {
        const key = normalizeExcelText(branch.locationName);
        if (!key || existing.has(key)) return;
        existing.add(key);
        additions.push({ id: makeId() + index + 70000, name: branch.locationName, city: branch.city || "To be updated", emirate: branch.emirate || "To be updated", type: "Mall" });
      });
      return additions.length ? [...additions, ...current] : current;
    });

    setGroups((current) => {
      let next = [...current];
      preview.importable.forEach((branch, index) => {
        const groupIndex = next.findIndex((group) => normalizeExcelText(group.name) === normalizeExcelText(branch.groupName));
        if (groupIndex === -1) {
          next.push({ id: makeId() + index + 80000, name: branch.groupName, concepts: branch.conceptName ? [branch.conceptName] : [], logoDataUrl: "", footerPhone: "", footerAddress: "", footerWebsite: "", footerColor: "#991b1e" });
          return;
        }
        if (branch.conceptName && !next[groupIndex].concepts?.some((concept) => normalizeExcelText(concept) === normalizeExcelText(branch.conceptName))) {
          next[groupIndex] = { ...next[groupIndex], concepts: [...(next[groupIndex].concepts || []), branch.conceptName] };
        }
      });
      return next;
    });

    setBranches((current) => {
      let next = [...current];
      const keyIndex = new Map(next.map((branch, index) => [normalizeBranchKey(branch.branchName, branch.locationName), index]));
      preview.parsedBranches.forEach((branch, index) => {
        const key = normalizeBranchKey(branch.branchName, branch.locationName);
        const existingIndex = keyIndex.get(key);
        if (existingIndex !== undefined && branchImport.duplicateMode === "skip") return;
        if (existingIndex !== undefined && branchImport.duplicateMode === "update") {
          next[existingIndex] = { ...next[existingIndex], ...branch, id: next[existingIndex].id, handover: next[existingIndex].handover || 0 };
          return;
        }
        next = [{ ...branch, id: makeId() + index + 90000 }, ...next];
      });
      return next;
    });

    alert(`Branch import completed.\n\nSaved/Updated: ${preview.importable.length}\nSkipped duplicates: ${preview.skipCount}\nMissing required rows: ${preview.missingRequired}`);
    resetBranchImport();
    closeModal();
    setTab("branches");
  }

  async function initializeVault(passphrase, confirmation) {
    if (!isAdmin) return { ok: false, message: "Only an administrator can create the credential vault." };
    if (!crypto?.subtle) return { ok: false, message: "Secure browser encryption is unavailable in this browser." };
    if (String(passphrase || "").length < 8) return { ok: false, message: "Use a vault passphrase with at least 8 characters." };
    if (passphrase !== confirmation) return { ok: false, message: "The passphrases do not match." };
    try {
      const salt = crypto.getRandomValues(new Uint8Array(16));
      const saltBase64 = bytesToBase64(salt);
      const key = await deriveVaultKey(passphrase, saltBase64);
      const verifier = await encryptVaultText("BRANCH_IT_CREDENTIAL_VAULT", key);
      setVaultSettings({ version: 1, salt: saltBase64, verifier, createdAt: new Date().toISOString() });
      vaultKeyRef.current = key;
      setVaultUnlocked(true);
      return { ok: true };
    } catch (error) {
      return { ok: false, message: `Unable to create the vault: ${error.message}` };
    }
  }

  async function unlockVault(passphrase) {
    if (!vaultSettings?.salt || !vaultSettings?.verifier) return { ok: false, message: "Vault configuration is missing." };
    try {
      const key = await deriveVaultKey(passphrase, vaultSettings.salt);
      const verifier = await decryptVaultText(vaultSettings.verifier, key);
      if (verifier !== "BRANCH_IT_CREDENTIAL_VAULT") throw new Error("Invalid passphrase");
      vaultKeyRef.current = key;
      setVaultUnlocked(true);
      return { ok: true };
    } catch {
      return { ok: false, message: "Incorrect vault passphrase." };
    }
  }

  function lockVault() {
    vaultKeyRef.current = null;
    setVaultUnlocked(false);
  }

  async function revealCredentialSecret(item) {
    if (!vaultKeyRef.current || !vaultUnlocked) throw new Error("Unlock the vault first.");
    return decryptVaultText(item.secretCipher, vaultKeyRef.current);
  }

  async function submitCredential(e) {
    e.preventDefault();
    if (!requireAdmin()) return;
    if (!vaultUnlocked || !vaultKeyRef.current) return setFormError("Unlock the Credential Vault before saving a record.");
    const f = normalizeCredentialRecord(forms.credential);
    const primaryName = f.type === "device" ? f.deviceName : f.type === "shared" ? (f.accountName || f.scopeName || f.platform) : f.owner;
    if (!primaryName) return setFormError("Please complete the main device, shared account or user name.");
    if (f.type === "access" && !f.linkedCredentialId) return setFormError("Please select the shared account assigned to this user.");
    if (f.type !== "access" && !f.username && !f.email) return setFormError("Please enter a username or login email address.");
    if (f.type === "device" && !f.branchId) return setFormError("Please select the branch for this device credential.");
    try {
      const existing = credentials.find((item) => item.id === editId);
      let secretCipher = existing?.secretCipher || "";
      if (f.type !== "access" && f.password) secretCipher = await encryptVaultText(f.password, vaultKeyRef.current);
      const item = {
        ...f,
        id: editId || makeId(),
        branchId: f.branchId ? Number(f.branchId) : "",
        linkedAssetId: f.linkedAssetId ? Number(f.linkedAssetId) : "",
        linkedCredentialId: f.linkedCredentialId ? Number(f.linkedCredentialId) : "",
        password: undefined,
        secretCipher: f.type === "access" ? "" : secretCipher,
        secretUpdatedAt: f.type !== "access" && f.password ? new Date().toISOString() : (existing?.secretUpdatedAt || ""),
        createdAt: existing?.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      delete item.password;
      setCredentials((items) => editId ? items.map((entry) => entry.id === editId ? item : entry) : [item, ...items]);
      closeModal();
      setForms((current) => ({ ...current, credential: emptyCredential }));
    } catch (error) {
      setFormError(`Unable to encrypt and save the password: ${error.message}`);
    }
  }

  function deleteCredential(item) {
    if (!requireAdmin()) return;
    const linkedAccess = item.type === "shared" ? credentials.filter((entry) => entry.type === "access" && Number(entry.linkedCredentialId) === Number(item.id)).length : 0;
    const linkedText = linkedAccess ? `\n\n${linkedAccess} access record(s) are linked to this shared account. They will remain but show as an unlinked account.` : "";
    askConfirm({ title: "Delete credential record?", message: `Delete ${item.deviceName || item.accountName || item.owner || item.username || "this credential"}? This cannot be undone.${linkedText}`, confirmText: "Delete Credential" }, () => setCredentials((items) => items.filter((entry) => entry.id !== item.id)));
  }

  function exportCredentialRegister() {
    const sharedAccounts = credentials.filter((item) => item.type === "shared");
    const rows = credentials.map((item) => {
      const branch = branches.find((branchItem) => branchItem.id === Number(item.branchId));
      const linkedAccount = sharedAccounts.find((entry) => Number(entry.id) === Number(item.linkedCredentialId));
      const linkedAsset = assets.find((asset) => Number(asset.id) === Number(item.linkedAssetId));
      return {
        Type: credentialTypeLabels[item.type] || item.type,
        Platform: item.platform || linkedAccount?.platform || "",
        Scope: item.scopeType || linkedAccount?.scopeType || "",
        "Scope Name": item.scopeName || linkedAccount?.scopeName || "",
        Branch: branch?.branchName || "",
        "Linked Shared Account": linkedAccount?.accountName || "",
        "Device / Account / User": item.deviceName || item.accountName || item.owner || "",
        "Linked Asset": linkedAsset?.deviceName || "",
        "Asset Serial": linkedAsset?.serialNumber || "",
        "IP Address": item.ipAddress || "",
        URL: item.url || "",
        Port: item.port || "",
        Username: item.username || "",
        Email: item.email || "",
        Mobile: item.mobile || "",
        "Account Owner": item.type === "shared" ? item.owner || "" : "",
        "Job Role": item.jobRole || "",
        "Department / Operation": item.departmentOperation || "",
        "MFA Status": item.mfaStatus || "",
        "Permission Level": item.permissionLevel || "",
        "Granted Date": item.grantedDate || "",
        "Approved By": item.approvedBy || "",
        "Last Password Change": item.lastPasswordChange || "",
        "Next Review Date": item.nextPasswordChange || "",
        "Revoked Date": item.revokedDate || "",
        Status: item.status || "",
        "Password Stored": item.secretCipher ? "Encrypted in vault" : "No",
        Notes: item.notes || "",
      };
    });
    if (!rows.length) return alert("No credential records to export.");
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = Object.keys(rows[0]).map((header) => ({ wch: Math.max(16, header.length + 2) }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Credential Register");
    XLSX.writeFile(workbook, `Credential_Vault_Register_${today()}.xlsx`);
  }

  async function importCredentialExcel(file) {
    if (!requireAdmin()) return;
    if (!vaultUnlocked || !vaultKeyRef.current) return alert("Unlock the Credential Vault before importing.");
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: false });
      const imported = [];
      const skipped = [];
      const norm = (value) => normalizeExcelText(value);
      const getValue = (row, aliases) => {
        const entries = Object.entries(row || {});
        for (const alias of aliases) {
          const found = entries.find(([key]) => norm(key) === norm(alias));
          if (found) return String(found[1] ?? "").trim();
        }
        return "";
      };
      for (const sheetName of workbook.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
        for (const row of rows) {
          const explicitType = getValue(row, ["Type", "Credential Type"]);
          const sheetKey = norm(sheetName);
          let type = norm(explicitType).includes("access") ? "access" : norm(explicitType).includes("shared") || norm(explicitType).includes("cloud") ? "shared" : norm(explicitType).includes("device") ? "device" : sheetKey.includes("access") ? "access" : (sheetKey.includes("shared") || sheetKey.includes("cloud") || sheetKey.includes("hik") || sheetKey.includes("connect")) ? "shared" : "device";
          const branchText = getValue(row, ["Branch", "Branch Access", "Location", "Site"]);
          const branch = branches.find((item) => norm(item.branchName) === norm(branchText) || norm(item.locationName) === norm(branchText));
          const linkedSerial = getValue(row, ["Linked Asset Serial", "Serial Number", "Serial No.", "Serial"]);
          const linkedAsset = linkedSerial ? assets.find((asset) => norm(asset.serialNumber) === norm(linkedSerial)) : null;
          const linkedAccountName = getValue(row, ["Linked Account", "Linked Shared Account", "Account Name"]);
          const sharedPool = [...credentials, ...imported].filter((item) => item.type === "shared");
          const linkedAccount = type === "access" ? sharedPool.find((item) => norm(item.accountName) === norm(linkedAccountName) || norm(`${item.platform} - ${item.scopeName}`) === norm(linkedAccountName)) : null;
          const password = type === "access" ? "" : getValue(row, ["Password", "New Password", "Current Password"]);
          const entry = normalizeCredentialRecord({
            id: makeId() + imported.length,
            type,
            branchId: branch?.id || linkedAsset?.branchId || "",
            linkedAssetId: linkedAsset?.id || "",
            linkedCredentialId: linkedAccount?.id || "",
            category: getValue(row, ["Category", "Device Type", "Credential Category"]) || (type === "device" ? "Other" : ""),
            deviceName: getValue(row, ["Device Name", "Device", "Asset Name", "NVR Name", "Server Name"]) || linkedAsset?.deviceName || "",
            platform: getValue(row, ["Platform", "Application", "Portal", "Service"]) || (sheetKey.includes("hik") ? "Hik-Connect" : ""),
            accountName: type === "shared" ? getValue(row, ["Account Name", "Account", "Profile Name", "User"]) : "",
            scopeType: getValue(row, ["Account Scope", "Scope Type", "Operation", "Department / Operation"]) || (type === "shared" ? "Other" : ""),
            scopeName: getValue(row, ["Scope Name", "Account Group", "User Group"]),
            ipAddress: getValue(row, ["IP Address", "IP", "Device IP"]),
            url: getValue(row, ["URL", "Portal URL", "Web Address"]),
            port: getValue(row, ["Port", "Service Port"]),
            username: type === "access" ? "" : getValue(row, ["Username", "User Name", "Login", "Login Name"]),
            email: getValue(row, ["Login Email", "Email", "E-mail", "Email ID"]),
            recoveryEmail: getValue(row, ["Recovery Email", "Backup Email"]),
            mobile: getValue(row, ["Recovery Mobile", "Mobile", "Phone", "Mobile Number"]),
            mfaStatus: getValue(row, ["MFA Status", "2FA", "Two Factor"]) || "Not Enabled",
            owner: type === "access" ? getValue(row, ["User Full Name", "Person Name", "User", "Owner"]) : getValue(row, ["Account Owner", "Owner"]),
            jobRole: getValue(row, ["Job Role", "Role / Position"]),
            departmentOperation: getValue(row, ["Department / Operation", "Department", "Operation"]),
            permissionLevel: getValue(row, ["Permission Level", "Access Level", "Role"]) || "Viewer",
            grantedDate: cleanDateText(getValue(row, ["Granted Date", "Access Date", "Created Date"])),
            approvedBy: getValue(row, ["Approved By", "Approver"]),
            revokedDate: cleanDateText(getValue(row, ["Revoked Date", "Removed Date"])),
            lastPasswordChange: cleanDateText(getValue(row, ["Last Password Change", "Last Password Changed", "Password Changed Date"])),
            nextPasswordChange: cleanDateText(getValue(row, ["Next Password Change", "Next Review Date", "Password Expiry", "Due Date"])),
            status: getValue(row, ["Status"]) || "Active",
            notes: getValue(row, ["Notes", "Remarks", "Comment"]),
            secretCipher: password ? await encryptVaultText(password, vaultKeyRef.current) : "",
            secretUpdatedAt: password ? new Date().toISOString() : "",
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
          });
          const mainName = entry.deviceName || entry.accountName || entry.owner || entry.username || entry.email;
          if (!mainName || (entry.type === "access" && !entry.linkedCredentialId)) { skipped.push(row); continue; }
          const duplicate = [...credentials, ...imported].some((existing) => {
            if (existing.type !== entry.type) return false;
            if (entry.type === "access") return Number(existing.linkedCredentialId) === Number(entry.linkedCredentialId) && norm(existing.owner) === norm(entry.owner);
            if (entry.type === "shared") return norm(existing.platform) === norm(entry.platform) && norm(existing.username || existing.email) === norm(entry.username || entry.email) && norm(existing.scopeType) === norm(entry.scopeType);
            return norm(existing.username || existing.email) === norm(entry.username || entry.email) && norm(existing.deviceName) === norm(entry.deviceName);
          });
          if (duplicate) { skipped.push(row); continue; }
          imported.push(entry);
        }
      }
      if (imported.length) setCredentials((items) => [...imported, ...items]);
      alert(`Credential import completed.\nImported: ${imported.length}\nSkipped/duplicates/unmatched access: ${skipped.length}`);
    } catch (error) {
      alert(`Credential import failed.\n\n${error.message}`);
    } finally {
      if (credentialImportRef.current) credentialImportRef.current.value = "";
    }
  }

  function submitLocation(e) {
    e.preventDefault();
    if (!requireAdmin()) return;
    const f = forms.location;
    if (!f.name || !f.city || !f.emirate) return setFormError("Please complete location name, city and emirate.");
    if (editId) {
      const oldLocation = locations.find((item) => item.id === editId);
      setLocations((current) => current.map((item) => (item.id === editId ? { ...item, ...f, id: editId } : item)));
      if (oldLocation) {
        setBranches((current) => current.map((branch) => branch.locationName === oldLocation.name ? { ...branch, locationName: f.name, city: f.city, emirate: f.emirate } : branch));
      }
    } else {
      setLocations((current) => [{ ...f, id: makeId() }, ...current]);
    }
    closeModal();
    setForms((x) => ({ ...x, location: emptyLocation }));
  }

  function deleteLocation(location) {
    if (!requireAdmin()) return;
    if (!location) return alert("Invalid location selected.");

    const locName = String(location.name || location.locationName || "").trim();
    const locCity = String(location.city || "").trim();
    const locEmirate = String(location.emirate || "").trim();

    const usedBranches = branches.filter((branch) => {
      if (!branch) return false;
      const branchLocationName = String(branch.locationName || "").trim();
      const branchCity = String(branch.city || "").trim();
      const branchEmirate = String(branch.emirate || "").trim();
      return branchLocationName === locName || normalizeLocationKey(branchLocationName, branchCity, branchEmirate) === normalizeLocationKey(locName, locCity, locEmirate);
    });

    if (usedBranches.length > 0) {
      const branchList = usedBranches.slice(0, 8).map((branch) => `- ${branch.branchName || "Unnamed Branch"}`).join("\n");
      const more = usedBranches.length > 8 ? `\n...and ${usedBranches.length - 8} more branch(es)` : "";
      alert(`Cannot delete location "${locName}" because it is used by ${usedBranches.length} branch(es).\n\n${branchList}${more}\n\nPlease edit or move those branches before deleting this location.`);
      return;
    }

    askConfirm({ title: "Delete Location?", message: `Location: ${locName}\nCity: ${locCity || "-"}\nEmirate: ${locEmirate || "-"}\n\nThis action cannot be undone.`, confirmText: "Delete Location" }, () => setLocations((current) => current.filter((item) => item && item.id !== location.id)));
  }

  function addAssetCategory(name, description = "") {
    const cleanName = String(name || "").trim();
    if (!cleanName) return { ok: false, message: "Enter a category name." };
    const existing = assetCategories.find((item) => normalizeExcelText(item.name) === normalizeExcelText(cleanName));
    if (existing) {
      if (existing.isActive === false) setAssetCategories((items) => items.map((item) => item.id === existing.id ? { ...item, isActive: true } : item));
      return { ok: true, category: { ...existing, isActive: true }, existing: true };
    }
    const category = { id: makeId(), name: cleanName, description: String(description || "").trim(), isActive: true };
    setAssetCategories((items) => [...items, category].sort((a, b) => a.name.localeCompare(b.name)));
    return { ok: true, category, existing: false };
  }

  function submitProduct(e) {
    e.preventDefault();
    if (!requireAdmin()) return;
    const f = forms.product;
    if (!f.productName || !f.category) return setFormError("Please complete Product Name and Category.");
    const item = normalizeAssetProduct({ ...f, id: editId || makeId(), productName: f.productName.trim(), category: f.category, brand: String(f.brand || "").trim(), model: String(f.model || "").trim(), description: f.description || "", defaultQuantity: f.defaultQuantity || "1", defaultLocation: f.defaultLocation || "", defaultStatus: f.defaultStatus || "Installed", remarks: cleanImportedPlaceholder(f.remarks) || "", createdAt: editId ? (assetProducts.find((product) => product.id === editId)?.createdAt || new Date().toISOString()) : new Date().toISOString(), updatedAt: new Date().toISOString() });
    const duplicate = assetProducts.find((product) => product.id !== editId && normalizeProductKey(product) === normalizeProductKey(item));
    if (duplicate) return setFormError(`A matching product already exists: ${duplicate.productName}. Edit the existing record instead.`);
    setAssetProducts((items) => editId ? items.map((product) => product.id === editId ? item : product) : [item, ...items]);
    closeModal();
    setForms((x) => ({ ...x, product: emptyAssetProduct }));
  }

  function submitAssetTemplate(e) {
    e.preventDefault();
    if (!requireAdmin()) return;
    const f = forms.assetTemplate;
    const rows = (Array.isArray(f.rows) ? f.rows : []).filter((row) => row.productId && Number(row.quantity || 0) > 0);
    if (!f.name || rows.length === 0) return setFormError("Please enter template name and add at least one product row.");
    const item = { id: editId || makeId(), name: f.name.trim(), branchType: f.branchType || "QSR", description: f.description || "", rows: rows.map((row) => ({ productId: row.productId, quantity: String(row.quantity || "1"), installedLocation: row.installedLocation || "", remarks: cleanImportedPlaceholder(row.remarks) || "" })) };
    setAssetTemplates((items) => editId ? items.map((template) => template.id === editId ? item : template) : [item, ...items]);
    closeModal();
    setForms((x) => ({ ...x, assetTemplate: emptyAssetTemplate }));
  }

  function applyBranchAssetTemplate(templateId) {
    if (!requireAdmin()) return;
    const template = assetTemplates.find((item) => String(item.id) === String(templateId));
    if (!template || !selectedBranch) return alert("Please select a valid branch asset template.");
    const created = [];
    (template.rows || []).forEach((row, index) => {
      const product = assetProducts.find((item) => String(item.id) === String(row.productId));
      if (!product) return;
      created.push({
        id: makeId() + index,
        branchId: selectedBranch.id,
        category: product.category || "Other IT Device",
        customCategory: "",
        deviceName: product.productName || "",
        description: product.description || "",
        quantity: String(row.quantity || product.defaultQuantity || "1"),
        brand: product.brand || "",
        model: product.model || "",
        serialNumber: "",
        ipAddress: "",
        macAddress: "",
        installedLocation: row.installedLocation || product.defaultLocation || "",
        switchPort: "",
        channelNumber: "",
        cameraView: "",
        warrantyStart: "",
        warrantyEnd: "",
        status: product.defaultStatus || "Installed",
        remarks: cleanImportedPlaceholder(row.remarks || product.remarks) || "",
        nvrUsername: product.category === "NVR" ? "admin" : "",
        nvrPassword: "",
        recordingDays: "",
        recordingStartDate: "",
        recordingEndDate: "",
        lastRecordingReviewDate: "",
        reviewedBy: "",
        reviewStatus: product.category === "NVR" ? "Review Pending" : "",
        issueFound: "",
        actionRequired: "",
        nextReviewDate: "",
      });
    });
    if (created.length === 0) return alert("No valid product rows found in this template.");
    askConfirm({ title: "Apply Branch Asset Template?", message: `Template: ${template.name}
Branch: ${selectedBranch.branchName}

This will add ${created.length} asset row(s). You can edit serial number, IP, MAC and warranty details after applying.`, confirmText: "Apply Template", tone: "normal" }, () => {
      setAssets((items) => [...created, ...items]);
      setBranchSubTab("assets");
      closeModal();
    });
  }

  function applyProductToAsset(product) { if (!product) return; updateForm("asset", "category", product.category || "Other IT Device"); updateForm("asset", "deviceName", product.productName || ""); updateForm("asset", "description", product.description || ""); updateForm("asset", "quantity", product.defaultQuantity || "1"); updateForm("asset", "brand", product.brand || ""); updateForm("asset", "model", product.model || ""); updateForm("asset", "installedLocation", product.defaultLocation || ""); updateForm("asset", "status", product.defaultStatus || "Installed"); if (product.remarks) updateForm("asset", "remarks", cleanImportedPlaceholder(product.remarks)); }

  function submitBranch(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.branch; if (!f.groupName || !f.branchName || !f.locationName || !f.openingDate || !f.completedDate) return setFormError("Please complete required branch fields."); const loc = locations.find((l) => l.name === f.locationName); const newB = { ...f, id: editId || makeId(), city: loc?.city || "To be updated", emirate: loc?.emirate || "To be updated", branchManager: f.branchManager || "To be updated", contactNumber: f.contactNumber || "To be updated", preparedBy: f.preparedBy || "Demo Technician", handover: editId ? selectedBranch.handover : 0, notes: f.notes || "New branch record created." }; if (editId) { setBranches((x) => x.map((b) => b.id === editId ? newB : b)); setSelectedBranch(newB); } else { setBranches((x) => [newB, ...x]); setNetworks((x) => [{ ...emptyNetwork, branchId: newB.id, isp: "Etisalat" }, ...x]); setSelectedBranch(newB); } closeModal(); setForms((x) => ({ ...x, branch: emptyBranch })); setTab("branch-detail"); }
  function submitAsset(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.asset; if (!f.category || !f.deviceName || !f.serialNumber || !f.installedLocation) return setFormError("Please complete category, device name, serial number and installed location."); const item = { ...f, id: editId || makeId(), branchId: selectedBranch.id, category: f.category === "Other IT Device" && f.customCategory ? f.customCategory : f.category, description: f.description || "", quantity: f.quantity || "1", brand: f.brand || "", model: f.model || "", ipAddress: f.ipAddress || "", macAddress: f.macAddress || "", switchPort: f.switchPort || "", warrantyStart: f.warrantyStart || "", warrantyEnd: f.warrantyEnd || "", remarks: cleanImportedPlaceholder(f.remarks) || "" }; setAssets((x) => editId ? x.map((a) => a.id === editId ? item : a) : [item, ...x]); closeModal(); setForms((x) => ({ ...x, asset: emptyAsset })); setBranchSubTab("assets"); }
  function submitNetwork(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.network; if (!f.isp || !f.routerIp) return setFormError("Please complete ISP and router IP."); const n = { ...f, branchId: selectedBranch.id, internetAccount: f.internetAccount || "To be updated", bandwidth: f.bandwidth || "To be updated", routerModel: f.routerModel || "To be updated", switchModel: f.switchModel || "To be updated", remarks: f.remarks || "No remarks" }; setNetworks((x) => x.some((v) => v.branchId === selectedBranch.id) ? x.map((v) => v.branchId === selectedBranch.id ? n : v) : [n, ...x]); closeModal(); }
  function submitFault(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.fault; if (!f.faultDate || !f.faultTime || !f.title || !f.attendedBy || !f.actionTaken) return setFormError("Please complete date, time, title, attended by and action taken."); const item = { ...f, id: editId || makeId(), branchId: Number(f.branchId || selectedBranch.id), description: f.description || "No description added.", reportedBy: f.reportedBy || "Not updated", closedDate: f.status === "Closed" ? (f.closedDate || f.faultDate) : "", remarks: f.remarks || "No remarks." }; setFaults((x) => editId ? x.map((v) => v.id === editId ? item : v) : [item, ...x]); closeModal(); setForms((x) => ({ ...x, fault: emptyFault })); }
  function submitHandover(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.handover; if (!f.handoverDate || !f.storeManagerName || !f.completedPersons || !f.inchargeName) return setFormError("Please complete handover date, store manager, completed persons and in-charge."); const item = { ...f, branchId: selectedBranch.id, operationsRemarks: f.operationsRemarks || "No operations remarks.", itRemarks: f.itRemarks || "No IT remarks." }; setHandovers((x) => x.some((h) => h.branchId === selectedBranch.id) ? x.map((h) => h.branchId === selectedBranch.id ? item : h) : [item, ...x]); if (item.handoverStatus === "Signed / Completed") { const updated = { ...selectedBranch, handover: 100, status: "Live" }; setBranches((x) => x.map((b) => b.id === updated.id ? updated : b)); setSelectedBranch(updated); } closeModal(); setBranchSubTab("handover"); }
  function submitPullout(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.pullout; if (!f.pulloutDate || !f.storeManagerName || !f.removedBy || !f.inchargeName || !f.destination) return setFormError("Please complete pullout date, manager, removed by, in-charge and destination."); const item = { ...f, branchId: selectedBranch.id, assetCondition: f.assetCondition || "Pending final asset inspection.", operationsRemarks: f.operationsRemarks || "No operations remarks.", itRemarks: f.itRemarks || "No IT remarks." }; setPullouts((x) => x.some((p) => p.branchId === selectedBranch.id) ? x.map((p) => p.branchId === selectedBranch.id ? item : p) : [item, ...x]); if (item.pulloutStatus === "Signed / Completed") { const updated = { ...selectedBranch, status: "Closed" }; setBranches((x) => x.map((b) => b.id === updated.id ? updated : b)); setSelectedBranch(updated); } closeModal(); setBranchSubTab("pullout"); }
  function submitProfile(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.profile; if (!f.name || !f.username || !f.roleTitle || !f.company) return setFormError("Please complete name, username, role and company."); const profile = { name: f.name, username: f.username.toLowerCase(), email: f.email || "", mobile: f.mobile || "", roleTitle: f.roleTitle, department: f.department || "IT Department", company: f.company, reportingTo: f.reportingTo || "To be updated", specialization: f.specialization || "To be updated", locationScope: f.locationScope || "Company-owned F&B branches across UAE", profileSummary: f.profileSummary || "Profile summary to be updated." }; setProfiles((x) => editId ? x.map((p) => p.username === editId ? profile : p) : [profile, ...x]); setUsers((x) => x.some((u) => u.username === profile.username) ? x.map((u) => u.username === profile.username ? { ...u, displayName: profile.name, role: f.accessLevel || u.role, password: f.password || u.password } : u) : [{ id: makeId(), username: profile.username, password: f.password || "ChangeMe@123", role: f.accessLevel || "user", displayName: profile.name }, ...x]); setSelectedContributor(profile.name); closeModal(); setForms((x) => ({ ...x, profile: emptyProfile })); }
  function submitGroupBranding(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.groupBranding; if (!f.name) return setFormError("Invalid group selected."); setGroups((items) => items.map((group) => group.id === f.id || group.name === f.name ? { ...group, ...f, footerColor: f.footerColor || "#991b1e" } : group)); closeModal(); }
  function submitCctv(e) { e.preventDefault(); if (!requireAdmin()) return; const f = forms.cctv; const calculatedDays = countRecordingDays(f.recordingStartDate, f.recordingEndDate); if (!f.nvrUsername || !f.recordingStartDate || !f.recordingEndDate || !f.lastRecordingReviewDate || !f.reviewedBy) return setFormError("Please complete username, recording start date, recording end date, review date and reviewed by."); if (!calculatedDays) return setFormError("Recording end date must be same as or after recording start date."); const autoStatus = f.reviewStatus && f.reviewStatus !== "Review Pending" ? f.reviewStatus : getAutoNvrStatus(calculatedDays); const nextReview = f.nextReviewDate || addDays(f.lastRecordingReviewDate || today(), 7); setAssets((x) => x.map((a) => a.id === editId ? { ...a, nvrUsername: f.nvrUsername, nvrPassword: f.nvrPassword || a.nvrPassword, recordingStartDate: f.recordingStartDate, recordingEndDate: f.recordingEndDate, recordingDays: calculatedDays, lastRecordingReviewDate: f.lastRecordingReviewDate, reviewedBy: f.reviewedBy, reviewStatus: autoStatus, issueFound: f.issueFound || "", actionRequired: f.actionRequired || "", nextReviewDate: nextReview, remarks: cleanImportedPlaceholder(f.remarks) || cleanImportedPlaceholder(a.remarks) } : a)); closeModal(); }

  function getProductUsage(product) {
    const templateCount = assetTemplates.filter((template) => (template.rows || []).some((row) => String(row.productId) === String(product.id))).length;
    const assetCount = assets.filter((asset) => normalizeExcelText(asset.category) === normalizeExcelText(product.category) && normalizeExcelText(asset.brand) === normalizeExcelText(product.brand) && normalizeExcelText(asset.model) === normalizeExcelText(product.model)).length;
    return { templateCount, assetCount, total: templateCount + assetCount };
  }

  function deleteProduct(id) {
    const product = assetProducts.find((item) => item.id === id);
    if (!requireAdmin() || !product) return;
    const usage = getProductUsage(product);
    if (usage.total > 0) {
      askConfirm({ title: "Deactivate Product?", message: `Product: ${product.productName}

This product is already used by ${usage.assetCount} asset record(s) and ${usage.templateCount} template(s). It cannot be permanently deleted. It will be marked Inactive and hidden from new entries.`, confirmText: "Deactivate Product", tone: "normal" }, () => setAssetProducts((items) => items.map((item) => item.id === id ? { ...item, isActive: false } : item)));
      return;
    }
    askConfirm({ title: "Delete Product?", message: `Product: ${product.productName}

This unused product will be permanently deleted.`, confirmText: "Delete Product" }, () => setAssetProducts((items) => items.filter((item) => item.id !== id)));
  }

  function setProductsActive(ids, isActive) {
    if (!requireAdmin() || !ids.length) return;
    setAssetProducts((items) => items.map((item) => ids.includes(item.id) ? { ...item, isActive } : item));
  }

  function deleteProductsBulk(ids) {
    if (!requireAdmin() || !ids.length) return;
    const selected = assetProducts.filter((item) => ids.includes(item.id));
    const used = selected.filter((item) => getProductUsage(item).total > 0);
    const unused = selected.filter((item) => getProductUsage(item).total === 0);
    askConfirm({ title: "Process Selected Products?", message: `${selected.length} product(s) selected.

${unused.length} unused product(s) will be permanently deleted.
${used.length} used product(s) will be deactivated to protect existing asset and template history.`, confirmText: "Continue" }, () => setAssetProducts((items) => items.filter((item) => !unused.some((product) => product.id === item.id)).map((item) => used.some((product) => product.id === item.id) ? { ...item, isActive: false } : item)));
  }

  function deleteAssetTemplate(id) {
    const template = assetTemplates.find((item) => item.id === id);
    if (!requireAdmin() || !template) return;
    askConfirm({ title: "Delete Branch Asset Template?", message: `Template: ${template.name}

This will not delete already saved branch assets.`, confirmText: "Delete Template" }, () => setAssetTemplates((items) => items.filter((item) => item.id !== id)));
  }

  function deleteBranch() {
    if (!requireAdmin()) return;
    if (branches.length <= 1) return alert("At least one branch record must remain.");
    const assetCount = assets.filter((a) => a.branchId === selectedBranch.id).length;
    const faultCount = faults.filter((f) => f.branchId === selectedBranch.id).length;
    askConfirm({ title: "Delete Branch?", message: `Branch: ${selectedBranch.branchName}

This will also remove ${assetCount} asset(s), network setup, handover, pullout and ${faultCount} fault log(s).

This action cannot be undone.`, confirmText: "Delete Branch" }, () => {
      const remaining = branches.filter((b) => b.id !== selectedBranch.id);
      setBranches(remaining);
      setAssets((x) => x.filter((a) => a.branchId !== selectedBranch.id));
      setNetworks((x) => x.filter((n) => n.branchId !== selectedBranch.id));
      setHandovers((x) => x.filter((h) => h.branchId !== selectedBranch.id));
      setPullouts((x) => x.filter((p) => p.branchId !== selectedBranch.id));
      setFaults((x) => x.filter((f) => f.branchId !== selectedBranch.id));
      setSelectedBranch(remaining[0]);
      setTab("branches");
    });
  }

  function deleteAsset(id) {
    const a = assets.find((x) => x.id === id);
    if (!requireAdmin() || !a) return;
    askConfirm({ title: "Delete IT Asset?", message: `Asset: ${a.deviceName}
Serial No: ${a.serialNumber || "No serial"}

This action cannot be undone.`, confirmText: "Delete Asset" }, () => setAssets((x) => x.filter((v) => v.id !== id)));
  }

  function deleteAssetsBulk(ids) {
    const list = Array.isArray(ids) ? ids : [];
    const selected = assets.filter((a) => list.includes(a.id));
    if (!requireAdmin() || selected.length === 0) return;
    const sample = selected.slice(0, 6).map((a) => `• ${a.deviceName} (${a.serialNumber || "No Serial"})`).join("\n");
    const more = selected.length > 6 ? `
...and ${selected.length - 6} more asset(s).` : "";
    askConfirm({ title: "Delete Selected Assets?", message: `Selected assets: ${selected.length}

${sample}${more}

This action cannot be undone.`, confirmText: "Delete Selected" }, () => setAssets((items) => items.filter((asset) => !list.includes(asset.id))));
  }

  function deleteFault(id) {
    if (!requireAdmin()) return;
    askConfirm({ title: "Delete Fault Log?", message: "This action cannot be undone.", confirmText: "Delete Fault" }, () => setFaults((x) => x.filter((v) => v.id !== id)));
  }

  function deleteUser(id) {
    if (!requireAdmin()) return;
    if (id === currentUser.id) return alert("You cannot delete the currently logged-in account.");
    askConfirm({ title: "Delete User Account?", message: "This action cannot be undone.", confirmText: "Delete User" }, () => setUsers((x) => x.filter((v) => v.id !== id)));
  }

  function cleanImportedExcelRemarks() {
    if (!requireAdmin()) return;
    const affected = assets.filter((asset) => asset.remarks && cleanImportedPlaceholder(asset.remarks) !== String(asset.remarks || "").trim());
    if (!affected.length) return alert("No imported Excel remark text found in current asset records.");
    const ok = confirm(`Clean imported Excel remark text from ${affected.length} asset record(s)?\n\nThis removes only auto remarks like Imported from Excel... and keeps manual remarks unchanged.`);
    if (!ok) return;
    setAssets((items) => items.map((asset) => ({ ...asset, remarks: cleanImportedPlaceholder(asset.remarks) })));
    alert(`Cleaned imported Excel remarks from ${affected.length} asset record(s).`);
  }

  function updateTemplateField(field, checked) {
    setTemplateBuilder((current) => ({ ...current, selected: { ...current.selected, [field]: checked } }));
  }

  function selectTemplatePreset(mode) {
    setTemplateBuilder((current) => {
      const def = templateDefinitions[current.type] || templateDefinitions.asset;
      const selected = {};
      def.fields.forEach((field) => {
        selected[field.field] = mode === "all" ? true : mode === "important" ? !!field.important : false;
      });
      return { ...current, selected };
    });
  }

  function downloadSelectedTemplate() {
    const def = templateDefinitions[templateBuilder.type] || templateDefinitions.asset;
    const chosen = def.fields.filter((field) => templateBuilder.selected?.[field.field]);
    if (!chosen.length) return alert("Select at least one field for the template.");
    const sample = {};
    chosen.forEach((field) => { sample[field.label] = def.sample?.[field.field] ?? ""; });
    const worksheet = XLSX.utils.json_to_sheet([sample], { header: chosen.map((field) => field.label) });
    worksheet["!cols"] = chosen.map((field) => ({ wch: Math.max(16, field.label.length + 4) }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, def.sheetName);
    XLSX.writeFile(workbook, `${def.fileName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function openModal(name, payload = null) {
    setFormError(""); setEditId(null);
    if (name === "branch") { const group = groups[0] || defaultGroups[0]; setForms((x) => ({ ...x, branch: { ...emptyBranch, groupName: group.name, conceptName: group.concepts[0] } })); }
    if (name === "editBranch") { setEditId(selectedBranch.id); setForms((x) => ({ ...x, branch: { ...emptyBranch, ...selectedBranch } })); name = "branch"; }
    if (name === "location") { setForms((x) => ({ ...x, location: emptyLocation })); if (payload) { setEditId(payload.id); setForms((x) => ({ ...x, location: { ...emptyLocation, ...payload } })); } }
    if (name === "asset") { setForms((x) => ({ ...x, asset: emptyAsset })); if (payload) { setEditId(payload.id); setForms((x) => ({ ...x, asset: { ...emptyAsset, ...payload, remarks: cleanImportedPlaceholder(payload.remarks) } })); } }
    if (name === "product") { setForms((x) => ({ ...x, product: emptyAssetProduct })); if (payload) { setEditId(payload.id); setForms((x) => ({ ...x, product: { ...normalizeAssetProduct(payload), remarks: cleanImportedPlaceholder(payload.remarks) } })); } }
    if (name === "assetTemplate") { setForms((x) => ({ ...x, assetTemplate: emptyAssetTemplate })); if (payload) { setEditId(payload.id); setForms((x) => ({ ...x, assetTemplate: { ...emptyAssetTemplate, ...payload, rows: Array.isArray(payload.rows) && payload.rows.length ? payload.rows : [{ ...emptyTemplateRow }] } })); } }
    if (name === "assetImport") resetExcelImport();
    if (name === "productImport") resetProductImport();
    if (name === "locationImport") resetLocationImport();
    if (name === "template") { const type = payload?.type || "asset"; setTemplateBuilder({ type, selected: defaultTemplateSelected(type) }); }
    if (name === "network") setForms((x) => ({ ...x, network: { ...emptyNetwork, ...(selectedNetwork || {}) } }));
    if (name === "fault") { setForms((x) => ({ ...x, fault: { ...emptyFault, branchId: String(selectedBranch.id), faultDate: today() } })); if (payload) { setEditId(payload.id); setForms((x) => ({ ...x, fault: { ...emptyFault, ...payload, branchId: String(payload.branchId) } })); } }
    if (name === "handover") setForms((x) => ({ ...x, handover: { ...emptyHandover, ...(selectedHandover || {}), handoverDate: selectedHandover?.handoverDate || selectedBranch.completedDate || today(), storeManagerName: selectedHandover?.storeManagerName || selectedBranch.branchManager || "" } }));
    if (name === "pullout") setForms((x) => ({ ...x, pullout: { ...emptyPullout, ...(selectedPullout || {}), pulloutDate: selectedPullout?.pulloutDate || today(), storeManagerName: selectedPullout?.storeManagerName || selectedBranch.branchManager || "" } }));
    if (name === "profile") { setForms((x) => ({ ...x, profile: emptyProfile })); if (payload) { setEditId(payload.username); const u = users.find((usr) => usr.username === payload.username); setForms((x) => ({ ...x, profile: { ...emptyProfile, ...payload, accessLevel: u?.role || "user", password: "" } })); } }
    if (name === "groupBranding") { if (!payload) return; setEditId(payload.id); setForms((x) => ({ ...x, groupBranding: { ...emptyGroupBranding, ...payload, footerColor: payload.footerColor || "#991b1e" } })); }
    if (name === "credential") {
      if (!vaultUnlocked || !vaultKeyRef.current) { alert("Unlock the Credential Vault first."); return; }
      setForms((current) => ({ ...current, credential: { ...emptyCredential, branchId: selectedBranch?.id ? String(selectedBranch.id) : "" } }));
      if (payload) {
        setEditId(payload.id);
        const normalizedCredential = normalizeCredentialRecord(payload);
        setForms((current) => ({ ...current, credential: { ...emptyCredential, ...normalizedCredential, branchId: normalizedCredential.branchId ? String(normalizedCredential.branchId) : "", linkedAssetId: normalizedCredential.linkedAssetId ? String(normalizedCredential.linkedAssetId) : "", linkedCredentialId: normalizedCredential.linkedCredentialId ? String(normalizedCredential.linkedCredentialId) : "", password: "" } }));
      }
    }
    if (name === "cctv") { const reviewDate = payload.lastRecordingReviewDate || today(); const nextDate = payload.nextReviewDate || addDays(reviewDate || today(), 7); const existingDays = Number(payload.recordingDays || 0); const recordingEndDate = payload.recordingEndDate || today(); const recordingStartDate = payload.recordingStartDate || (existingDays > 0 ? addDays(recordingEndDate, -(existingDays - 1)) : ""); const calculatedDays = countRecordingDays(recordingStartDate, recordingEndDate) || existingDays || ""; setEditId(payload.id); setForms((x) => ({ ...x, cctv: { nvrUsername: payload.nvrUsername || "admin", nvrPassword: payload.nvrPassword || "", recordingStartDate, recordingEndDate, recordingDays: String(calculatedDays || ""), lastRecordingReviewDate: reviewDate, reviewedBy: payload.reviewedBy || "Demo Technician", reviewStatus: payload.reviewStatus || getAutoNvrStatus(calculatedDays), issueFound: payload.issueFound || "", actionRequired: payload.actionRequired || "", nextReviewDate: nextDate, remarks: cleanImportedPlaceholder(payload.remarks) || "" } })); }
    setModal(name);
  }

  function downloadFaultCSV() {
    const rows = filteredFaults.filter((f) => (!reportRange.start || f.faultDate >= reportRange.start) && (!reportRange.end || f.faultDate <= reportRange.end));
    if (!rows.length) return alert("No fault logs found for selected range.");
    const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const header = ["Fault Date", "Fault Time", "Branch", "Location", "Category", "Title", "Reported By", "Attended By", "Action Taken", "Status", "Closed Date", "Remarks"];
    const csv = [header, ...rows.map((f) => [f.faultDate, f.faultTime, f.branch.branchName, f.branch.locationName, f.category, f.title, f.reportedBy, f.attendedBy, f.actionTaken, f.status, f.closedDate, f.remarks])].map((r) => r.map(esc).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `fault-log-report-${reportRange.start || "all"}-to-${reportRange.end || "all"}.csv`; a.click(); URL.revokeObjectURL(a.href);
  }
  function printReport(type) { document.title = `${selectedBranch.branchName} - ${type} Report`; setTimeout(() => window.print(), 200); }

  if (!currentUser) {
    return (
      <div className="login-v842 min-h-screen overflow-hidden bg-[#0D1317] text-[#E7ECEE]">
        <style>{`
          @keyframes login-spin{to{transform:rotate(360deg)}}
          @keyframes login-pop{0%{transform:scale(.96);opacity:.65}100%{transform:scale(1);opacity:1}}
          @keyframes login-line{0%,100%{opacity:.28}50%{opacity:.7}}
          .login-v842 input,.login-v842 textarea{color:#E7ECEE}
          .login-v842 input::placeholder,.login-v842 textarea::placeholder{color:#687780}
          .login-network-line{animation:login-line 4.5s ease-in-out infinite}
          .login-success-pop{animation:login-pop .28s ease-out both}
        `}</style>
        <div className="mx-auto flex min-h-screen max-w-[1400px] items-center justify-center p-4 sm:p-6 lg:p-8">
          <div className="grid w-full max-w-[1150px] overflow-hidden rounded-[20px] border border-[#2A343B] bg-[#141B21] shadow-[0_35px_100px_rgba(0,0,0,.42)] lg:grid-cols-2">
            <section className="relative flex min-h-[580px] flex-col overflow-hidden border-b border-[#2A343B] p-7 sm:p-10 lg:border-b-0 lg:border-r lg:p-12">
              <svg className="pointer-events-none absolute inset-0 h-full w-full opacity-60" viewBox="0 0 600 620" fill="none" aria-hidden="true">
                <path className="login-network-line" d="M62 84 193 210 375 160 530 254" stroke="#C98A4B" strokeOpacity=".42" strokeWidth="1.4" />
                <path className="login-network-line" d="M193 210 108 382 264 552 385 290 530 438" stroke="#C98A4B" strokeOpacity=".34" strokeWidth="1.2" style={{animationDelay:"1.2s"}} />
                <path d="M375 160 385 290 530 254" stroke="#6C93D9" strokeOpacity=".18" strokeWidth="1" />
                {[{x:62,y:84,r:5,c:'#4FAE9B'},{x:193,y:210,r:7,c:'#C98A4B'},{x:375,y:160,r:6,c:'#C98A4B'},{x:108,y:382,r:5,c:'#6C93D9'},{x:264,y:552,r:8,c:'#C98A4B'},{x:385,y:290,r:5,c:'#4FAE9B'},{x:530,y:438,r:6,c:'#C98A4B'}].map((dot,index)=><circle key={index} cx={dot.x} cy={dot.y} r={dot.r} fill={dot.c} fillOpacity=".72" />)}
              </svg>

              <div className="relative z-10">
                <div className="flex items-center gap-3 font-mono text-[11px] uppercase tracking-[0.22em] text-[#C98A4B] sm:text-xs">
                  <span className="h-2 w-2 rounded-full bg-[#4FAE9B] shadow-[0_0_12px_#4FAE9B]" />
                  <span>{branches.length} Branches</span><span>·</span><span>{locations.length} Locations Online</span>
                </div>
                <h1 className="mt-5 font-serif text-4xl font-bold tracking-tight text-white sm:text-[42px]">Branch IT Control</h1>
                <p className="mt-4 max-w-xl text-base leading-7 text-[#9CADB6] sm:text-lg">Secure access for IT assets, CCTV, network setup, and branch handover records across every location.</p>

                <div className="mt-9 space-y-3">
                  <div className="rounded-xl border border-[#2A343B] bg-[#1B242B]/95 p-4 shadow-sm backdrop-blur-sm">
                    <div className="flex items-start gap-3">
                      <span className="mt-0.5 rounded-md bg-[rgba(201,138,75,.16)] px-2 py-1 font-mono text-[10px] font-bold text-[#C98A4B]">ADMIN</span>
                      <div><h3 className="text-lg font-semibold text-white">Full access</h3><p className="mt-1 text-sm leading-6 text-[#8A98A0]">Add, edit, delete, export or restore the database, and update operational records.</p></div>
                    </div>
                  </div>
                  <div className="rounded-xl border border-[#2A343B] bg-[#1B242B]/95 p-4 shadow-sm backdrop-blur-sm">
                    <div className="flex items-start gap-3">
                      <span className="mt-0.5 rounded-md bg-[rgba(79,174,155,.16)] px-2 py-1 font-mono text-[10px] font-bold text-[#4FAE9B]">USER</span>
                      <div><h3 className="text-lg font-semibold text-white">View only</h3><p className="mt-1 text-sm leading-6 text-[#8A98A0]">Search records and preview reports with controlled, read-only access.</p></div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="relative z-10 mt-auto border-t border-[#2A343B] pt-5">
                <p className="font-mono text-[9px] font-medium uppercase tracking-[0.25em] text-[#687780]">Powered By</p>
                <p className="mt-1.5 text-xs font-semibold tracking-wide text-[#E7ECEE]">Crescent IT Solution</p>
              </div>
            </section>

            <section className="flex min-h-[580px] items-center bg-[#11191E] p-7 sm:p-10 lg:p-12">
              <div className="mx-auto w-full max-w-md">
                <h2 className="text-3xl font-bold tracking-tight text-white">Sign in</h2>
                <p className="mt-2 text-base text-[#8A98A0]">Enter your credentials to continue.</p>

                {loginError && <div className="mt-5 rounded-xl border border-[rgba(217,106,93,.45)] bg-[rgba(217,106,93,.12)] p-4 text-sm font-semibold text-[#F19A90]">{loginError}</div>}

                <form onSubmit={login} className="mt-8 space-y-5">
                  <label className="block">
                    <span className="text-xs font-bold uppercase tracking-wide text-[#8FA2AD]">Username</span>
                    <input autoComplete="username" value={loginForm.username} onChange={(e)=>setLoginForm((x)=>({...x,username:e.target.value}))} placeholder="Enter your username" className="mt-2 h-14 w-full rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 text-base outline-none transition focus:border-[#C98A4B] focus:ring-4 focus:ring-[rgba(201,138,75,.12)]" />
                  </label>
                  <label className="block">
                    <div className="flex items-center justify-between gap-3"><span className="text-xs font-bold uppercase tracking-wide text-[#8FA2AD]">Password</span><button type="button" onClick={openForgotPassword} className="text-xs font-semibold text-[#C98A4B] transition hover:text-[#E2A35F]">Forgot password?</button></div>
                    <input autoComplete="current-password" type="password" value={loginForm.password} onChange={(e)=>setLoginForm((x)=>({...x,password:e.target.value}))} placeholder="Enter your password" className="mt-2 h-14 w-full rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 text-base outline-none transition focus:border-[#C98A4B] focus:ring-4 focus:ring-[rgba(201,138,75,.12)]" />
                  </label>
                  <button type="submit" disabled={loginStatus === "loading" || loginStatus === "success"} className={`flex h-14 w-full items-center justify-center gap-3 rounded-xl px-5 text-sm font-bold transition duration-200 ${loginStatus === "success" ? "bg-[#4FAE9B] text-[#07110F]" : "bg-[#D4914D] text-[#17110A] hover:-translate-y-0.5 hover:bg-[#E0A15F] hover:shadow-[0_12px_30px_rgba(201,138,75,.22)] active:translate-y-0 disabled:cursor-wait disabled:opacity-90"}`}>
                    {loginStatus === "loading" && <span className="h-5 w-5 rounded-full border-2 border-[#17110A]/35 border-t-[#17110A]" style={{animation:"login-spin .75s linear infinite"}} />}
                    {loginStatus === "success" && <span className="login-success-pop text-lg">✓</span>}
                    <span>{loginStatus === "loading" ? "Signing in..." : loginStatus === "success" ? "Access granted" : "Sign in"}</span>
                    {loginStatus === "idle" && <span aria-hidden="true">→</span>}
                  </button>
                </form>

                <div className="mt-7 border-t border-[#2A343B] pt-6">
                  <p className="text-sm font-semibold text-white">Authorized access only.</p>
                  <p className="mt-1.5 text-sm leading-6 text-[#8A98A0]">Contact the system administrator if you need account access or password support.</p>
                </div>
              </div>
            </section>
          </div>
        </div>

        {forgotOpen && <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/70 p-4 backdrop-blur-sm" onMouseDown={(event)=>{if(event.target===event.currentTarget) closeForgotPassword();}}>
          <div className="w-full max-w-lg rounded-2xl border border-[#2A343B] bg-[#141B21] p-6 shadow-[0_30px_80px_rgba(0,0,0,.55)] sm:p-7">
            <div className="flex items-start justify-between gap-4"><div><p className="font-mono text-[10px] uppercase tracking-[0.2em] text-[#C98A4B]">Account Recovery</p><h3 className="mt-2 text-2xl font-bold text-white">Forgot password</h3><p className="mt-2 text-sm leading-6 text-[#8A98A0]">Submit a request. An administrator will verify your account and set a temporary password.</p></div><button type="button" onClick={closeForgotPassword} className="flex h-9 w-9 items-center justify-center rounded-lg border border-[#2A343B] bg-[#1B242B] text-[#8A98A0] hover:text-white">×</button></div>
            {forgotSuccess ? <div className="mt-6 rounded-xl border border-[rgba(79,174,155,.4)] bg-[rgba(79,174,155,.12)] p-5"><p className="font-semibold text-[#79CDBB]">Request submitted</p><p className="mt-2 text-sm leading-6 text-[#A9B8BE]">{forgotSuccess}</p><button type="button" onClick={closeForgotPassword} className="mt-5 w-full rounded-xl bg-[#D4914D] px-5 py-3 text-sm font-bold text-[#17110A]">Close</button></div> : <form onSubmit={submitForgotPassword} className="mt-6 space-y-4">
              {forgotError && <div className="rounded-xl border border-[rgba(217,106,93,.45)] bg-[rgba(217,106,93,.12)] p-3 text-sm font-semibold text-[#F19A90]">{forgotError}</div>}
              <label className="block"><span className="text-xs font-bold uppercase tracking-wide text-[#8FA2AD]">Username *</span><input value={forgotForm.username} onChange={(e)=>setForgotForm((x)=>({...x,username:e.target.value}))} className="mt-2 h-12 w-full rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 outline-none focus:border-[#C98A4B]" placeholder="Your login username" /></label>
              <label className="block"><span className="text-xs font-bold uppercase tracking-wide text-[#8FA2AD]">Registered email or mobile *</span><input value={forgotForm.contact} onChange={(e)=>setForgotForm((x)=>({...x,contact:e.target.value}))} className="mt-2 h-12 w-full rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 outline-none focus:border-[#C98A4B]" placeholder="Email address or mobile number" /></label>
              <label className="block"><span className="text-xs font-bold uppercase tracking-wide text-[#8FA2AD]">Message</span><textarea value={forgotForm.message} onChange={(e)=>setForgotForm((x)=>({...x,message:e.target.value}))} rows="3" className="mt-2 w-full rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 py-3 outline-none focus:border-[#C98A4B]" placeholder="Optional information for the administrator" /></label>
              <div className="flex gap-3"><button type="button" onClick={closeForgotPassword} className="flex-1 rounded-xl border border-[#2A343B] bg-[#1B242B] px-4 py-3 text-sm font-semibold text-[#C4CED3]">Cancel</button><button type="submit" disabled={forgotSubmitting} className="flex flex-1 items-center justify-center gap-2 rounded-xl bg-[#D4914D] px-4 py-3 text-sm font-bold text-[#17110A] disabled:opacity-70">{forgotSubmitting && <span className="h-4 w-4 rounded-full border-2 border-[#17110A]/35 border-t-[#17110A]" style={{animation:"login-spin .75s linear infinite"}} />}{forgotSubmitting ? "Submitting..." : "Submit request"}</button></div>
            </form>}
          </div>
        </div>}
      </div>
    );
  }

  return <div data-theme={theme} className={`app-shell min-h-screen overflow-x-hidden ${theme === "dark" ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900"}`}>
    <style>{`
      .app-shell[data-theme="dark"] .bg-white{background-color:#0f172a!important;color:#e2e8f0!important}.app-shell[data-theme="dark"] .bg-slate-50{background-color:#111827!important;color:#e2e8f0!important}.app-shell[data-theme="dark"] .bg-slate-100{background-color:#1e293b!important;color:#e2e8f0!important}.app-shell[data-theme="dark"] .border-slate-200,.app-shell[data-theme="dark"] .border-slate-300{border-color:#334155!important}.app-shell[data-theme="dark"] .text-slate-900,.app-shell[data-theme="dark"] .text-slate-800,.app-shell[data-theme="dark"] .text-slate-700,.app-shell[data-theme="dark"] .text-slate-600{color:#e2e8f0!important}.app-shell[data-theme="dark"] .text-slate-500,.app-shell[data-theme="dark"] .text-slate-400{color:#94a3b8!important}.app-shell[data-theme="dark"] input,.app-shell[data-theme="dark"] select,.app-shell[data-theme="dark"] textarea{background-color:#0f172a!important;color:#e2e8f0!important;border-color:#475569!important}.app-shell[data-theme="dark"] .printable-report,.app-shell[data-theme="dark"] .printable-report *{background-color:white!important;color:#020617!important}
      .app-shell{background:#f5f7fb!important}.professional-sidebar{background:linear-gradient(160deg,#020617 0%,#0f172a 58%,#172554 100%);box-shadow:18px 0 45px rgba(2,6,23,.18)}.professional-main{background:linear-gradient(180deg,#f8fafc 0%,#f5f7fb 52%,#eef2f7 100%)}.professional-topbar{background:rgba(248,250,252,.95);backdrop-filter:blur(16px);box-shadow:0 1px 0 rgba(148,163,184,.22),0 8px 28px rgba(15,23,42,.035)}.professional-page-card{border:1px solid rgba(219,228,240,.95);background:#fff;box-shadow:0 10px 30px rgba(15,23,42,.06)}.professional-sidebar nav{scrollbar-width:thin;scrollbar-color:rgba(148,163,184,.28) transparent}.professional-sidebar nav::-webkit-scrollbar{width:5px}.professional-sidebar nav::-webkit-scrollbar-thumb{background:rgba(148,163,184,.28);border-radius:999px}.professional-sidebar nav::-webkit-scrollbar-track{background:transparent}.app-shell button{transition:all .18s ease}.app-shell table thead{background:#f8fafc!important}.app-shell table tbody tr{transition:background-color .15s ease}.app-shell input,.app-shell select,.app-shell textarea{transition:border-color .15s ease,box-shadow .15s ease}.app-shell input:focus,.app-shell select:focus,.app-shell textarea:focus{border-color:#2563eb!important;box-shadow:0 0 0 3px rgba(37,99,235,.10)}.app-shell[data-theme="dark"]{background:#020617!important}.app-shell[data-theme="dark"] .professional-topbar{background:rgba(15,23,42,.92)}.app-shell[data-theme="dark"] .professional-page-card{border-color:#334155;background:#0f172a}.dashboard-donut{position:relative;border-radius:9999px}.dashboard-donut:after{content:"";position:absolute;inset:24%;border-radius:9999px;background:#fff}.app-shell[data-theme="dark"] .dashboard-donut:after{background:#0f172a}.dashboard-donut-label{position:absolute;inset:0;z-index:2;display:flex;align-items:center;justify-content:center;flex-direction:column}.scope-card-active{border-color:#2563eb!important;background:linear-gradient(145deg,#eff6ff,#ffffff)!important;box-shadow:0 12px 30px rgba(37,99,235,.13)!important}.sidebar-nav{overflow:hidden!important}.professional-sidebar{overscroll-behavior:none}@media(min-width:1024px) and (max-height:720px){.professional-sidebar{padding-top:8px!important;padding-bottom:8px!important}.sidebar-nav button{padding-top:6px!important;padding-bottom:6px!important;font-size:12px!important}.sidebar-nav{margin-top:0!important}.professional-sidebar .mt-auto{gap:5px!important}.professional-sidebar .mt-auto>div{padding-top:6px!important;padding-bottom:6px!important}}
      @media(min-width:1024px){.professional-main{width:calc(100% - 18rem)!important}.app-shell .stat-card{min-height:132px}.app-shell .professional-page-card{border-radius:16px}.app-shell .dashboard-donut-label .text-3xl{font-size:1.5rem!important}.app-shell .dashboard-donut-label .text-\[11px\]{font-size:10px!important}.app-shell table{font-size:12px}.app-shell table th,.app-shell table td{padding-top:9px!important;padding-bottom:9px!important}}
      @media(max-width:767px){html,body{max-width:100%!important;overflow-x:hidden!important}.app-shell{width:100%!important;max-width:100vw!important;overflow-x:hidden!important}.app-shell *{box-sizing:border-box!important;min-width:0}.app-shell main{width:100%!important;max-width:100vw!important;overflow-x:hidden!important;padding:8px!important}.app-shell section{width:100%!important;max-width:100%!important;overflow-x:hidden!important;margin-top:10px!important}.app-shell input,.app-shell select,.app-shell textarea,.app-shell button{max-width:100%!important}.app-shell .rounded-2xl{border-radius:13px!important}.app-shell .p-8{padding:14px!important}.app-shell .p-6{padding:12px!important}.app-shell .p-5{padding:10px!important}.app-shell .p-4{padding:9px!important}.app-shell .gap-4{gap:9px!important}.app-shell h2.text-2xl,.app-shell h2.md\\:text-3xl{font-size:20px!important;line-height:1.1!important}.app-shell h3.text-2xl{font-size:18px!important}.app-shell .text-sm{font-size:12px!important;line-height:1.3!important}.app-shell .text-xs{font-size:10.5px!important}.app-shell .dashboard-mobile-grid,.app-shell .asset-summary-mobile-grid{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px!important}.app-shell .dashboard-mobile-grid .stat-card{min-height:92px!important;padding:10px!important}.app-shell .dashboard-mobile-grid .stat-card h3{font-size:22px!important;line-height:1!important;margin-top:7px!important}.app-shell .dashboard-project-mobile-grid{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px!important}.app-shell .dashboard-project-mobile-grid .dashboard-project-card{padding:10px!important;min-height:120px!important}.app-shell .dashboard-project-mobile-grid .dashboard-project-card:nth-child(3){grid-column:span 2/span 2}.app-shell .mobile-top-card{padding:9px!important;margin-bottom:8px!important}.app-shell .mobile-top-card h1{font-size:15px!important}.app-shell .mobile-menu select{height:40px!important;font-size:12px!important}.app-shell .mobile-record-card{position:relative;overflow:hidden;padding:12px!important;border-radius:18px!important;border:1px solid #e2e8f0!important;background:linear-gradient(180deg,#fff 0%,#f8fafc 100%)!important;box-shadow:0 8px 20px rgba(15,23,42,.07)!important}.app-shell .mobile-record-card:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:#0f172a}.app-shell .mobile-record-header{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;padding-left:4px}.app-shell .mobile-record-kicker{margin:0 0 3px!important;color:#64748b!important;font-size:10px!important;font-weight:800!important;text-transform:uppercase;letter-spacing:.04em}.app-shell .mobile-record-title{color:#020617!important;font-size:14px!important;font-weight:900!important;line-height:1.2!important;overflow-wrap:anywhere!important}.app-shell .mobile-record-subtitle{margin-top:4px!important;color:#475569!important;font-size:11.5px!important;line-height:1.25!important;font-weight:600!important;overflow-wrap:anywhere!important}.app-shell .mobile-record-status{flex-shrink:0;max-width:36%;text-align:right}.app-shell .mobile-record-status .inline-flex{padding:4px 9px!important;font-size:10px!important;line-height:1!important;white-space:nowrap}.app-shell .mobile-record-details{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin-top:12px;padding-left:4px}.app-shell .mobile-detail-tile{min-height:58px;border:1px solid #e2e8f0;border-radius:14px;background:rgba(255,255,255,.82);padding:8px}.app-shell .mobile-cell-label{color:#64748b!important;font-size:9.5px!important;font-weight:800!important;text-transform:uppercase;letter-spacing:.03em;margin:0 0 4px!important}.app-shell .mobile-cell-value{color:#0f172a!important;font-size:12px!important;line-height:1.25!important;font-weight:750!important;overflow-wrap:anywhere!important}.app-shell .mobile-record-actions{margin-top:10px;padding-left:4px;display:flex;justify-content:flex-end}.app-shell .mobile-record-actions .flex{width:100%;display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-end}.app-shell .mobile-record-actions button{flex:1 1 auto;min-width:72px;padding:8px 10px!important;font-size:11px!important;border-radius:12px!important}}
      /* Refined charcoal dark mode based on the approved dashboard reference */
      .app-shell[data-theme="dark"]{
        --dark-bg:#0D1317;
        --dark-surface:#141B21;
        --dark-surface-2:#1B242B;
        --dark-border:#2A343B;
        --dark-text:#E7ECEE;
        --dark-muted:#8A98A0;
        --dark-accent:#C98A4B;
        --dark-accent-soft:rgba(201,138,75,.16);
        --dark-good:#4FAE9B;
        --dark-good-soft:rgba(79,174,155,.16);
        --dark-danger:#D96A5D;
        --dark-danger-soft:rgba(217,106,93,.16);
        --dark-blue:#6C93D9;
        --dark-blue-soft:rgba(108,147,217,.16);
        --dark-purple:#9B7FD4;
        color-scheme:dark;
        background:var(--dark-bg)!important;
        color:var(--dark-text)!important;
      }
      .app-shell[data-theme="dark"] .professional-main{background:var(--dark-bg)!important}
      .app-shell[data-theme="dark"] .professional-sidebar{background:var(--dark-surface)!important;border-right:1px solid var(--dark-border)!important;box-shadow:none!important}
      .app-shell[data-theme="dark"] .professional-topbar{background:rgba(20,27,33,.96)!important;border-color:var(--dark-border)!important;box-shadow:0 1px 0 var(--dark-border)!important}
      .app-shell[data-theme="dark"] .professional-page-card,
      .app-shell[data-theme="dark"] .bg-white{background-color:var(--dark-surface)!important;color:var(--dark-text)!important}
      .app-shell[data-theme="dark"] .bg-slate-50,
      .app-shell[data-theme="dark"] .bg-slate-100,
      .app-shell[data-theme="dark"] .bg-slate-200,
      .app-shell[data-theme="dark"] .bg-slate-800,
      .app-shell[data-theme="dark"] .bg-slate-900,
      .app-shell[data-theme="dark"] .bg-slate-950{background-color:var(--dark-surface-2)!important;color:var(--dark-text)!important}
      .app-shell[data-theme="dark"] [class*="from-blue-50"][class*="via-white"][class*="to-indigo-50"]{background:var(--dark-surface)!important;border-color:var(--dark-border)!important;box-shadow:none!important}
      .app-shell[data-theme="dark"] .border-slate-100,
      .app-shell[data-theme="dark"] .border-slate-200,
      .app-shell[data-theme="dark"] .border-slate-300,
      .app-shell[data-theme="dark"] .divide-slate-100>*+*,
      .app-shell[data-theme="dark"] .divide-slate-200>*+*{border-color:var(--dark-border)!important}
      .app-shell[data-theme="dark"] .text-slate-950,
      .app-shell[data-theme="dark"] .text-slate-900,
      .app-shell[data-theme="dark"] .text-slate-800,
      .app-shell[data-theme="dark"] .text-slate-700,
      .app-shell[data-theme="dark"] .text-slate-600,
      .app-shell[data-theme="dark"] .text-slate-200,
      .app-shell[data-theme="dark"] .text-slate-100{color:var(--dark-text)!important}
      .app-shell[data-theme="dark"] .text-slate-500,
      .app-shell[data-theme="dark"] .text-slate-400,
      .app-shell[data-theme="dark"] .text-slate-300{color:var(--dark-muted)!important}
      .app-shell[data-theme="dark"] .text-blue-600,
      .app-shell[data-theme="dark"] .text-blue-700{color:var(--dark-blue)!important}
      .app-shell[data-theme="dark"] button.text-blue-600,
      .app-shell[data-theme="dark"] .professional-topbar .text-blue-600{color:var(--dark-accent)!important}
      .app-shell[data-theme="dark"] .bg-blue-50,
      .app-shell[data-theme="dark"] .bg-blue-100{background-color:var(--dark-blue-soft)!important}
      .app-shell[data-theme="dark"] .bg-cyan-50,
      .app-shell[data-theme="dark"] .bg-cyan-100,
      .app-shell[data-theme="dark"] .bg-teal-50,
      .app-shell[data-theme="dark"] .bg-teal-100{background-color:var(--dark-good-soft)!important}
      .app-shell[data-theme="dark"] .text-cyan-600,
      .app-shell[data-theme="dark"] .text-cyan-700,
      .app-shell[data-theme="dark"] .text-teal-600,
      .app-shell[data-theme="dark"] .text-teal-700{color:var(--dark-good)!important}
      .app-shell[data-theme="dark"] .ring-cyan-100,
      .app-shell[data-theme="dark"] .ring-teal-100{--tw-ring-color:rgba(79,174,155,.34)!important}
      .app-shell[data-theme="dark"] .stat-card-active{border-color:var(--dark-accent)!important;box-shadow:0 0 0 2px rgba(201,138,75,.18)!important}
      .app-shell[data-theme="dark"] .stat-card-icon[data-accent="teal"]{background:var(--dark-good-soft)!important;color:var(--dark-good)!important}
      .app-shell[data-theme="dark"] .donut-track{stroke:var(--dark-border)!important}
      .app-shell[data-theme="dark"] .border-blue-100,
      .app-shell[data-theme="dark"] .border-blue-200{border-color:rgba(108,147,217,.34)!important}
      .app-shell[data-theme="dark"] .bg-emerald-50,
      .app-shell[data-theme="dark"] .bg-emerald-100,
      .app-shell[data-theme="dark"] .bg-green-50,
      .app-shell[data-theme="dark"] .bg-green-100{background-color:var(--dark-good-soft)!important}
      .app-shell[data-theme="dark"] .text-emerald-600,
      .app-shell[data-theme="dark"] .text-emerald-700,
      .app-shell[data-theme="dark"] .text-green-600,
      .app-shell[data-theme="dark"] .text-green-700{color:var(--dark-good)!important}
      .app-shell[data-theme="dark"] .border-emerald-200,
      .app-shell[data-theme="dark"] .border-green-200{border-color:rgba(79,174,155,.34)!important}
      .app-shell[data-theme="dark"] .bg-amber-50,
      .app-shell[data-theme="dark"] .bg-amber-100,
      .app-shell[data-theme="dark"] .bg-yellow-50,
      .app-shell[data-theme="dark"] .bg-yellow-100{background-color:var(--dark-accent-soft)!important}
      .app-shell[data-theme="dark"] .text-amber-600,
      .app-shell[data-theme="dark"] .text-amber-700,
      .app-shell[data-theme="dark"] .text-yellow-600,
      .app-shell[data-theme="dark"] .text-yellow-700{color:var(--dark-accent)!important}
      .app-shell[data-theme="dark"] .border-amber-200,
      .app-shell[data-theme="dark"] .border-yellow-200{border-color:rgba(201,138,75,.34)!important}
      .app-shell[data-theme="dark"] .bg-violet-50,
      .app-shell[data-theme="dark"] .bg-violet-100,
      .app-shell[data-theme="dark"] .bg-purple-50,
      .app-shell[data-theme="dark"] .bg-purple-100,
      .app-shell[data-theme="dark"] .bg-indigo-50,
      .app-shell[data-theme="dark"] .bg-indigo-100{background-color:rgba(155,127,212,.16)!important}
      .app-shell[data-theme="dark"] .text-violet-600,
      .app-shell[data-theme="dark"] .text-violet-700,
      .app-shell[data-theme="dark"] .text-purple-600,
      .app-shell[data-theme="dark"] .text-purple-700,
      .app-shell[data-theme="dark"] .text-indigo-600,
      .app-shell[data-theme="dark"] .text-indigo-700{color:var(--dark-purple)!important}
      .app-shell[data-theme="dark"] .border-violet-200,
      .app-shell[data-theme="dark"] .border-purple-200,
      .app-shell[data-theme="dark"] .border-indigo-200{border-color:rgba(155,127,212,.34)!important}
      .app-shell[data-theme="dark"] .bg-red-50,
      .app-shell[data-theme="dark"] .bg-red-100{background-color:var(--dark-danger-soft)!important}
      .app-shell[data-theme="dark"] .text-red-600,
      .app-shell[data-theme="dark"] .text-red-700,
      .app-shell[data-theme="dark"] .text-red-800{color:var(--dark-danger)!important}
      .app-shell[data-theme="dark"] .border-red-200,
      .app-shell[data-theme="dark"] .border-red-300{border-color:rgba(217,106,93,.34)!important}
      .app-shell[data-theme="dark"] .bg-blue-600{background-color:var(--dark-accent)!important;color:#11181d!important}
      .app-shell[data-theme="dark"] .sidebar-nav-active{background:var(--dark-surface-2)!important;border-color:var(--dark-border)!important;color:var(--dark-text)!important;box-shadow:none!important}
      .app-shell[data-theme="dark"] .sidebar-nav-active svg{color:var(--dark-accent)!important}
      .app-shell[data-theme="dark"] .scope-card-active{border-color:var(--dark-accent)!important;background:var(--dark-surface)!important;box-shadow:0 0 0 1px rgba(201,138,75,.18)!important}
      .app-shell[data-theme="dark"] .dashboard-donut:after{background:var(--dark-surface)!important}
      .app-shell[data-theme="dark"] table thead{background:var(--dark-surface-2)!important;color:var(--dark-muted)!important}
      .app-shell[data-theme="dark"] table tbody tr{border-color:var(--dark-border)!important}
      .app-shell[data-theme="dark"] .hover\:bg-slate-50:hover,
      .app-shell[data-theme="dark"] table tbody tr:hover{background:var(--dark-surface-2)!important}
      .app-shell[data-theme="dark"] input,
      .app-shell[data-theme="dark"] select,
      .app-shell[data-theme="dark"] textarea{background:var(--dark-surface-2)!important;color:var(--dark-text)!important;border-color:var(--dark-border)!important}
      .app-shell[data-theme="dark"] input::placeholder,
      .app-shell[data-theme="dark"] textarea::placeholder{color:var(--dark-muted)!important}
      .app-shell[data-theme="dark"] input:focus,
      .app-shell[data-theme="dark"] select:focus,
      .app-shell[data-theme="dark"] textarea:focus{border-color:var(--dark-accent)!important;box-shadow:0 0 0 3px rgba(201,138,75,.12)!important}
      .app-shell[data-theme="dark"] .shadow-sm,
      .app-shell[data-theme="dark"] .shadow-md,
      .app-shell[data-theme="dark"] .shadow-lg,
      .app-shell[data-theme="dark"] .shadow-xl{box-shadow:none!important}
      .app-shell[data-theme="dark"] .mobile-record-card{background:var(--dark-surface)!important;border-color:var(--dark-border)!important;box-shadow:none!important}
      .app-shell[data-theme="dark"] .mobile-record-card:before{background:var(--dark-accent)!important}
      .app-shell[data-theme="dark"] .mobile-detail-tile{background:var(--dark-surface-2)!important;border-color:var(--dark-border)!important}
      .app-shell[data-theme="dark"] .mobile-record-title,
      .app-shell[data-theme="dark"] .mobile-cell-value{color:var(--dark-text)!important}
      .app-shell[data-theme="dark"] .mobile-record-kicker,
      .app-shell[data-theme="dark"] .mobile-record-subtitle,
      .app-shell[data-theme="dark"] .mobile-cell-label{color:var(--dark-muted)!important}
      .app-shell[data-theme="dark"] .printable-report,
      .app-shell[data-theme="dark"] .printable-report *{background-color:white!important;color:#020617!important}

      /* v8.40 — exact charcoal theme and interaction polish from approved reference */
      .app-shell[data-theme="dark"]{
        --ref-bg:#0D1317;
        --ref-surface:#141B21;
        --ref-surface-2:#1B242B;
        --ref-surface-hover:#222D35;
        --ref-border:#2A343B;
        --ref-text:#E7ECEE;
        --ref-muted:#8A98A0;
        --ref-accent:#C98A4B;
        --ref-accent-hover:#D89A5A;
        --ref-accent-soft:rgba(201,138,75,.16);
        --ref-good:#4FAE9B;
        --ref-good-soft:rgba(79,174,155,.16);
        --ref-danger:#D96A5D;
        --ref-danger-soft:rgba(217,106,93,.16);
        --ref-blue:#6C93D9;
        --ref-blue-soft:rgba(108,147,217,.16);
        --ref-purple:#9B7FD4;
        background:var(--ref-bg)!important;
        color:var(--ref-text)!important;
        font-family:Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif!important;
      }
      .app-shell[data-theme="dark"] h1,
      .app-shell[data-theme="dark"] h2,
      .app-shell[data-theme="dark"] h3,
      .app-shell[data-theme="dark"] .font-bold{
        font-family:"Space Grotesk",Inter,ui-sans-serif,system-ui,sans-serif;
      }
      .app-shell[data-theme="dark"] .professional-main{
        background:var(--ref-bg)!important;
      }
      .app-shell[data-theme="dark"] .professional-sidebar{
        background:var(--ref-surface)!important;
        border-right:1px solid var(--ref-border)!important;
        box-shadow:none!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] .professional-topbar{
        min-height:64px!important;
        padding-left:32px!important;
        padding-right:32px!important;
        background:rgba(13,19,23,.98)!important;
        border-color:var(--ref-border)!important;
        box-shadow:none!important;
        backdrop-filter:none!important;
      }
      .app-shell[data-theme="dark"] .professional-topbar > div:last-child > div,
      .app-shell[data-theme="dark"] .professional-topbar > div:last-child > button{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-border)!important;
        box-shadow:none!important;
      }
      .app-shell[data-theme="dark"] .professional-topbar > div:last-child > button:hover{
        background:var(--ref-surface-hover)!important;
        border-color:rgba(201,138,75,.5)!important;
        color:var(--ref-accent)!important;
      }
      .app-shell[data-theme="dark"] .professional-page-card,
      .app-shell[data-theme="dark"] .stat-card,
      .app-shell[data-theme="dark"] .bg-white{
        background:var(--ref-surface)!important;
        border-color:var(--ref-border)!important;
        color:var(--ref-text)!important;
        box-shadow:none!important;
      }
      .app-shell[data-theme="dark"] .rounded-3xl{border-radius:14px!important}
      .app-shell[data-theme="dark"] .rounded-2xl{border-radius:12px!important}
      .app-shell[data-theme="dark"] .rounded-xl{border-radius:9px!important}
      .app-shell[data-theme="dark"] .bg-slate-50,
      .app-shell[data-theme="dark"] .bg-slate-100,
      .app-shell[data-theme="dark"] .bg-slate-200,
      .app-shell[data-theme="dark"] .bg-slate-800,
      .app-shell[data-theme="dark"] .bg-slate-900,
      .app-shell[data-theme="dark"] .bg-slate-950{
        background:var(--ref-surface-2)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] .professional-page-card:hover,
      .app-shell[data-theme="dark"] .stat-card:hover{
        background:var(--ref-surface)!important;
        border-color:var(--ref-border)!important;
        box-shadow:none!important;
        transform:none!important;
      }
      .app-shell[data-theme="dark"] .text-slate-950,
      .app-shell[data-theme="dark"] .text-slate-900,
      .app-shell[data-theme="dark"] .text-slate-800,
      .app-shell[data-theme="dark"] .text-slate-700,
      .app-shell[data-theme="dark"] .text-slate-600,
      .app-shell[data-theme="dark"] .text-slate-200,
      .app-shell[data-theme="dark"] .text-slate-100,
      .app-shell[data-theme="dark"] .text-white{
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] .text-slate-500,
      .app-shell[data-theme="dark"] .text-slate-400,
      .app-shell[data-theme="dark"] .text-slate-300{
        color:var(--ref-muted)!important;
      }
      .app-shell[data-theme="dark"] .border-slate-100,
      .app-shell[data-theme="dark"] .border-slate-200,
      .app-shell[data-theme="dark"] .border-slate-300,
      .app-shell[data-theme="dark"] .border-slate-400,
      .app-shell[data-theme="dark"] .divide-slate-100>*+*,
      .app-shell[data-theme="dark"] .divide-slate-200>*+*{
        border-color:var(--ref-border)!important;
      }
      .app-shell[data-theme="dark"] .sidebar-nav-active{
        background:var(--ref-surface-2)!important;
        border-color:transparent!important;
        color:var(--ref-text)!important;
        box-shadow:none!important;
      }
      .app-shell[data-theme="dark"] .sidebar-nav-active svg{color:var(--ref-accent)!important}
      .app-shell[data-theme="dark"] .sidebar-nav button:not(.sidebar-nav-active){
        color:var(--ref-muted)!important;
        background:transparent!important;
        border-color:transparent!important;
      }
      .app-shell[data-theme="dark"] .sidebar-nav button:not(.sidebar-nav-active):hover{
        color:var(--ref-text)!important;
        background:var(--ref-surface-2)!important;
        border-color:transparent!important;
        transform:none!important;
      }
      .app-shell[data-theme="dark"] .professional-sidebar [class*="bg-white/"],
      .app-shell[data-theme="dark"] .professional-sidebar [class*="bg-slate-950/"]{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-border)!important;
      }
      .app-shell[data-theme="dark"] .professional-sidebar .bg-indigo-600,
      .app-shell[data-theme="dark"] .professional-sidebar .bg-indigo-700{
        background:var(--ref-blue-soft)!important;
        color:var(--ref-blue)!important;
      }
      .app-shell[data-theme="dark"] button{
        box-shadow:none!important;
      }
      .app-shell[data-theme="dark"] button:hover{
        transform:none!important;
      }
      .app-shell[data-theme="dark"] button.bg-slate-900,
      .app-shell[data-theme="dark"] button.bg-slate-950,
      .app-shell[data-theme="dark"] button.bg-blue-600,
      .app-shell[data-theme="dark"] button.bg-blue-700,
      .app-shell[data-theme="dark"] button.bg-indigo-600,
      .app-shell[data-theme="dark"] button.bg-indigo-700{
        background:var(--ref-accent)!important;
        border-color:var(--ref-accent)!important;
        color:#11181D!important;
      }
      .app-shell[data-theme="dark"] button.bg-slate-900:hover,
      .app-shell[data-theme="dark"] button.bg-slate-950:hover,
      .app-shell[data-theme="dark"] button.bg-blue-600:hover,
      .app-shell[data-theme="dark"] button.bg-blue-700:hover,
      .app-shell[data-theme="dark"] button.bg-indigo-600:hover,
      .app-shell[data-theme="dark"] button.bg-indigo-700:hover{
        background:var(--ref-accent-hover)!important;
        border-color:var(--ref-accent-hover)!important;
        color:#11181D!important;
      }
      .app-shell[data-theme="dark"] button.bg-white,
      .app-shell[data-theme="dark"] button[class*="border-slate"]{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-border)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] button.bg-white:hover,
      .app-shell[data-theme="dark"] button[class*="border-slate"]:hover{
        background:var(--ref-surface-hover)!important;
        border-color:rgba(201,138,75,.55)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] .hover\\:bg-slate-50:hover,
      .app-shell[data-theme="dark"] .hover\\:bg-slate-100:hover,
      .app-shell[data-theme="dark"] .hover\\:bg-slate-200:hover,
      .app-shell[data-theme="dark"] .hover\\:bg-white:hover,
      .app-shell[data-theme="dark"] .active\\:bg-slate-50:active{
        background:var(--ref-surface-hover)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] .hover\\:border-blue-200:hover,
      .app-shell[data-theme="dark"] .hover\\:border-slate-300:hover{
        border-color:rgba(201,138,75,.55)!important;
      }
      .app-shell[data-theme="dark"] .hover\\:text-blue-600:hover,
      .app-shell[data-theme="dark"] .hover\\:text-blue-700:hover{
        color:var(--ref-accent)!important;
      }
      .app-shell[data-theme="dark"] a,
      .app-shell[data-theme="dark"] button.text-blue-600,
      .app-shell[data-theme="dark"] button.text-blue-700{
        color:var(--ref-accent)!important;
      }
      .app-shell[data-theme="dark"] table{
        background:var(--ref-surface)!important;
        border-color:var(--ref-border)!important;
      }
      .app-shell[data-theme="dark"] table thead,
      .app-shell[data-theme="dark"] table thead.bg-slate-100{
        background:var(--ref-surface-2)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] table th,
      .app-shell[data-theme="dark"] table td{
        border-color:var(--ref-border)!important;
      }
      .app-shell[data-theme="dark"] table tbody tr,
      .app-shell[data-theme="dark"] table tbody tr.bg-white,
      .app-shell[data-theme="dark"] table tbody tr.bg-slate-50,
      .app-shell[data-theme="dark"] table tbody tr.bg-yellow-50,
      .app-shell[data-theme="dark"] table tbody tr.bg-red-50{
        background:var(--ref-surface)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] table tbody tr:hover,
      .app-shell[data-theme="dark"] table tbody tr.hover\\:bg-slate-50:hover,
      .app-shell[data-theme="dark"] table tbody tr.hover\\:bg-yellow-100:hover,
      .app-shell[data-theme="dark"] table tbody tr.hover\\:bg-red-100:hover{
        background:var(--ref-surface-hover)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] table tbody tr:hover td,
      .app-shell[data-theme="dark"] table tbody tr:hover p,
      .app-shell[data-theme="dark"] table tbody tr:hover span{
        color:inherit;
      }
      .app-shell[data-theme="dark"] input,
      .app-shell[data-theme="dark"] select,
      .app-shell[data-theme="dark"] textarea{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-border)!important;
        color:var(--ref-text)!important;
      }
      .app-shell[data-theme="dark"] input::placeholder,
      .app-shell[data-theme="dark"] textarea::placeholder{color:var(--ref-muted)!important}
      .app-shell[data-theme="dark"] input:focus,
      .app-shell[data-theme="dark"] select:focus,
      .app-shell[data-theme="dark"] textarea:focus{
        border-color:var(--ref-accent)!important;
        box-shadow:0 0 0 3px rgba(201,138,75,.13)!important;
      }
      .app-shell[data-theme="dark"] .scope-card-active{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-accent)!important;
        box-shadow:0 0 0 1px rgba(201,138,75,.16)!important;
      }
      .app-shell[data-theme="dark"] .dashboard-donut:after{background:var(--ref-surface)!important}
      .app-shell[data-theme="dark"] .shadow-sm,
      .app-shell[data-theme="dark"] .shadow-md,
      .app-shell[data-theme="dark"] .shadow-lg,
      .app-shell[data-theme="dark"] .shadow-xl,
      .app-shell[data-theme="dark"] [class*="shadow-["]{box-shadow:none!important}
      .app-shell[data-theme="dark"] .mobile-record-card{
        background:var(--ref-surface)!important;
        border-color:var(--ref-border)!important;
        box-shadow:none!important;
      }
      .app-shell[data-theme="dark"] .mobile-detail-tile{
        background:var(--ref-surface-2)!important;
        border-color:var(--ref-border)!important;
      }
      .app-shell[data-theme="dark"] .mobile-record-card:before{background:var(--ref-accent)!important}
      @media(min-width:1024px){
        .app-shell[data-theme="dark"] .professional-sidebar{width:250px!important;padding:16px!important}
        .app-shell[data-theme="dark"] .professional-main{margin-left:250px!important;width:calc(100% - 250px)!important}
        .app-shell[data-theme="dark"] .professional-main > div:last-child{padding-left:32px!important;padding-right:32px!important;padding-top:28px!important}
        .app-shell[data-theme="dark"] .stat-card{min-height:0!important;padding:16px!important}
        .app-shell[data-theme="dark"] .professional-page-card{border-radius:12px!important}
      }
      /* v8.47 modal and form contrast upgrade */
      .app-modal-overlay{
        background:rgba(2,6,23,.78)!important;
        backdrop-filter:blur(3px);
        -webkit-backdrop-filter:blur(3px);
      }
      .app-modal-shell,
      .app-confirm-shell{
        position:relative;
        background:#ffffff;
        color:#0f172a;
        border:2px solid #cbd5e1;
        box-shadow:0 28px 90px rgba(2,6,23,.38),0 0 0 1px rgba(255,255,255,.55);
      }
      .app-modal-header{
        background:#f8fafc;
        border-bottom:1px solid #cbd5e1;
      }
      .app-modal-body{background:#ffffff}
      .app-modal-close{
        background:#e2e8f0;
        color:#334155;
        border:1px solid #cbd5e1;
      }
      .app-modal-close:hover{background:#cbd5e1;color:#0f172a}
      .app-modal-shell form{
        border:1px solid #dbe3ec;
        border-radius:16px;
        background:#f8fafc;
        padding:18px;
      }
      .app-modal-shell form input,
      .app-modal-shell form select,
      .app-modal-shell form textarea{
        background:#ffffff!important;
        border-width:1.5px!important;
        border-color:#94a3b8!important;
        box-shadow:inset 0 1px 2px rgba(15,23,42,.035);
      }
      .app-modal-shell form input:hover,
      .app-modal-shell form select:hover,
      .app-modal-shell form textarea:hover{border-color:#64748b!important}
      .app-modal-shell form input:focus,
      .app-modal-shell form select:focus,
      .app-modal-shell form textarea:focus{
        border-color:#2563eb!important;
        box-shadow:0 0 0 3px rgba(37,99,235,.13)!important;
      }
      .app-modal-shell form label>span{color:#334155!important}
      .app-modal-actions{
        border-top:1px solid #cbd5e1;
        background:transparent;
      }
      .app-confirm-shell .app-modal-actions{background:#f8fafc;border-radius:0 0 22px 22px}
      .app-shell[data-theme="dark"] .app-modal-overlay{
        background:rgba(2,6,10,.82)!important;
      }
      .app-shell[data-theme="dark"] .app-modal-shell,
      .app-shell[data-theme="dark"] .app-confirm-shell{
        background:#182128!important;
        color:#E7ECEE!important;
        border-color:#C98A4B!important;
        box-shadow:0 30px 100px rgba(0,0,0,.62),0 0 0 1px rgba(201,138,75,.22)!important;
      }
      .app-shell[data-theme="dark"] .app-modal-header{
        background:#202b33!important;
        border-color:#3b4851!important;
      }
      .app-shell[data-theme="dark"] .app-modal-body{background:#182128!important}
      .app-shell[data-theme="dark"] .app-modal-close{
        background:#2a353d!important;
        color:#E7ECEE!important;
        border-color:#46545e!important;
      }
      .app-shell[data-theme="dark"] .app-modal-close:hover{
        background:#C98A4B!important;
        color:#11181d!important;
        border-color:#C98A4B!important;
      }
      .app-shell[data-theme="dark"] .app-modal-shell form{
        background:#11191e!important;
        border-color:#3b4851!important;
        box-shadow:inset 0 0 0 1px rgba(255,255,255,.018);
      }
      .app-shell[data-theme="dark"] .app-modal-shell form input,
      .app-shell[data-theme="dark"] .app-modal-shell form select,
      .app-shell[data-theme="dark"] .app-modal-shell form textarea{
        background:#202b33!important;
        color:#E7ECEE!important;
        border-color:#55636d!important;
        box-shadow:inset 0 1px 2px rgba(0,0,0,.26)!important;
      }
      .app-shell[data-theme="dark"] .app-modal-shell form input:hover,
      .app-shell[data-theme="dark"] .app-modal-shell form select:hover,
      .app-shell[data-theme="dark"] .app-modal-shell form textarea:hover{border-color:#778690!important}
      .app-shell[data-theme="dark"] .app-modal-shell form input:focus,
      .app-shell[data-theme="dark"] .app-modal-shell form select:focus,
      .app-shell[data-theme="dark"] .app-modal-shell form textarea:focus{
        border-color:#C98A4B!important;
        box-shadow:0 0 0 3px rgba(201,138,75,.18)!important;
      }
      .app-shell[data-theme="dark"] .app-modal-shell form label>span{color:#E7ECEE!important}
      .app-shell[data-theme="dark"] .app-modal-actions{border-color:#3b4851!important}
      .app-shell[data-theme="dark"] .app-confirm-shell .app-modal-actions{background:#11191e!important}
      @media(max-width:640px){
        .app-modal-overlay{padding:8px!important}
        .app-modal-header{padding:16px!important}
        .app-modal-body{padding-left:16px!important;padding-right:16px!important;padding-bottom:16px!important}
        .app-modal-shell form{padding:14px!important}
      }
      @media print{.company-handover{box-shadow:none!important;border-radius:0!important}.company-handover table{page-break-inside:auto}.company-handover tr{page-break-inside:avoid;page-break-after:auto}@page{size:A4 portrait;margin:0}body{margin:0!important;background:white!important}body *{visibility:hidden!important}.printable-report,.printable-report *,.store-print-pages,.store-print-pages *{visibility:visible!important}.printable-report{position:absolute!important;left:0!important;top:0!important;width:198mm!important;margin:0!important;padding:6mm!important;box-shadow:none!important;border:none!important;background:white!important;color:#020617!important;font-size:10px!important;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}.store-handover-screen{display:none!important}.store-print-pages{display:block!important;position:absolute!important;left:0!important;top:0!important;width:210mm!important;margin:0!important;background:white!important;color:#020617!important;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}.store-print-page{position:relative!important;width:210mm!important;min-height:297mm!important;height:297mm!important;margin:0!important;padding:8mm 10mm 30mm 10mm!important;background:white!important;page-break-after:always!important;break-after:page!important;overflow:hidden!important}.store-print-page:last-child{page-break-after:auto!important;break-after:auto!important}.store-print-header{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:8mm!important;border-bottom:1.1mm solid #0f172a!important;padding-bottom:3.5mm!important;margin-bottom:5mm!important}.store-print-logo{width:43mm!important;height:19mm!important;border:0.25mm solid #cbd5e1!important;border-radius:3mm!important;display:flex!important;align-items:center!important;justify-content:center!important;padding:2mm!important}.store-print-logo img{max-width:100%!important;max-height:100%!important;object-fit:contain!important}.store-print-company{text-align:right!important;font-size:9px!important;line-height:1.35!important;color:#334155!important}.store-print-title{text-align:center!important;margin:4mm 0 5mm!important}.store-print-title h2{font-size:18px!important;line-height:1.2!important;margin:0!important;text-transform:uppercase!important;font-weight:900!important}.store-print-title p{font-size:9px!important;color:#64748b!important;margin:1.5mm 0 0!important}.store-print-section{margin-top:4.5mm!important;break-inside:avoid!important}.store-print-section h3{font-size:12.5px!important;line-height:1.22!important;margin:0 0 3mm!important;font-weight:900!important}.store-print-table{width:100%!important;border-collapse:collapse!important;font-size:9px!important;line-height:1.28!important}.store-print-table th{background:#0f172a!important;color:white!important;font-weight:800!important}.store-print-table th,.store-print-table td{border:0.25mm solid #dbe4f0!important;padding:1.9mm!important;vertical-align:top!important}.store-print-table tbody tr:nth-child(odd){background:#f8fafc!important}.store-print-info td{font-size:9px!important;padding:1.9mm!important}.store-print-text{font-size:11.2px!important;line-height:1.58!important}.store-print-text p{margin:0 0 2.6mm!important}.store-print-text ul,.store-print-text ol{margin:0!important;padding-left:5mm!important}.store-print-text li{margin-bottom:2.4mm!important}.store-print-signatures{display:grid!important;grid-template-columns:1fr 1fr!important;gap:8mm!important;margin-top:7mm!important}.store-print-signbox{border:0.3mm solid #cbd5e1!important;border-radius:3.5mm!important;padding:7mm!important;min-height:72mm!important;font-size:12px!important;line-height:1.85!important}.store-print-signbox strong{display:block!important;font-size:12.5px!important;margin-bottom:3mm!important}.store-print-signbox p{margin:3mm 0!important;font-size:12px!important;line-height:1.7!important}.store-print-footer{position:absolute!important;left:10mm!important;right:10mm!important;bottom:6mm!important;height:16mm!important;display:grid!important;grid-template-columns:1fr 1.7fr 1fr!important;align-items:center!important;gap:4mm!important;padding:2.8mm 6mm!important;color:white!important;font-size:8.2px!important;font-weight:800!important}.store-print-footer *{color:white!important}.store-handover-print-footer,.store-handover-flow-footer,.store-handover-print-spacer{display:none!important}.no-print{display:none!important}.pdf-card{border:1px solid #cbd5e1!important;border-radius:8px!important;padding:8px!important;break-inside:avoid!important}.pdf-table th,.pdf-table td{border-bottom:1px solid #e2e8f0!important;padding:4px!important;font-size:8px!important}.pdf-row{display:grid!important;grid-template-columns:1fr 1fr 1fr!important;gap:6px!important}.pdf-two{display:grid!important;grid-template-columns:1fr 1fr!important;gap:8px!important}.pdf-signatures{display:grid!important;grid-template-columns:1fr 1fr 1fr!important;gap:8px!important}}
    `}</style>
    <div className="min-h-screen w-full min-w-0 overflow-x-hidden">
      <aside className="professional-sidebar fixed inset-y-0 left-0 z-40 hidden h-screen w-72 flex-col overflow-hidden px-3 py-3 text-white lg:flex">
        <div className="flex shrink-0 items-center gap-2.5 px-2 py-1.5">
          <div className="flex h-9 w-9 items-center justify-center rounded-xl border border-white/15 bg-white/10 text-blue-100 shadow-lg"><AppIcon name="handover" className="h-5 w-5" /></div>
          <div><h1 className="text-sm font-bold tracking-wide">BRANCH IT CONTROL</h1><p className="mt-0.5 text-[10px] text-blue-100/70">Asset & Handover System</p></div>
        </div>
        <div className="my-2.5 h-px shrink-0 bg-white/10" />
        <nav className="sidebar-nav shrink-0 space-y-0.5 overflow-hidden">{menuItems.map(([key, label, icon]) => <button key={key} onClick={() => setTab(key)} className={`flex w-full items-center gap-2.5 rounded-lg border px-3 py-2 text-left text-[13px] font-semibold leading-5 transition ${tab === key ? "sidebar-nav-active border-blue-400/25 bg-[#1e3a8a] text-white shadow-lg shadow-slate-950/25" : "border-transparent text-slate-200 hover:border-white/10 hover:bg-white/10 hover:text-white"}`}><AppIcon name={icon} className="h-4 w-4" /><span>{label}</span></button>)}</nav>
        <div className="mt-auto shrink-0 space-y-2 pt-2">
          <div className="rounded-xl border border-white/10 bg-white/[0.07] px-3 py-2.5">
            <div className="flex items-center justify-between gap-2"><p className="text-[9px] font-bold uppercase tracking-[0.16em] text-blue-100/60">Selected Branch</p><div className="flex items-center gap-1.5"><span className="h-1.5 w-1.5 rounded-full bg-emerald-400" /><span className="text-[10px] text-slate-300">{selectedBranch.status}</span></div></div>
            <p className="mt-1 truncate text-xs font-semibold">{selectedBranch.branchName}</p><p className="mt-0.5 truncate text-[10px] text-slate-300">{selectedBranch.locationName}</p>
          </div>
          <div className="flex items-center gap-2.5 rounded-xl border border-white/10 bg-slate-950/20 px-3 py-2">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-indigo-600 text-xs font-bold">{String(currentUser.displayName || "U").slice(0,1).toUpperCase()}</div>
            <div className="min-w-0 flex-1"><p className="truncate text-xs font-semibold">{currentUser.displayName}</p><p className="text-[10px] text-slate-300">{isAdmin ? "Super Admin" : "View Only"}</p></div>
            <button onClick={logout} title="Logout" className="rounded-lg p-1.5 text-slate-300 hover:bg-white/10 hover:text-white"><AppIcon name="logout" className="h-3.5 w-3.5" /></button>
          </div>
          <div className="rounded-xl border border-white/15 bg-white/[0.06] px-3 py-2"><p className="text-[8px] font-bold uppercase tracking-[0.18em] text-slate-400">Powered By</p><p className="mt-1 text-[11px] font-semibold text-white">Crescent IT Solution</p></div>
        </div>
      </aside>
      <main className="professional-main min-h-screen min-w-0 max-w-full overflow-x-hidden p-3 sm:p-4 lg:ml-72 lg:w-[calc(100%-18rem)] lg:p-0">
        <header className="professional-topbar sticky top-0 z-20 hidden min-h-[74px] items-center justify-between border-b border-slate-200/80 px-5 lg:flex">
          <div><p className="text-sm text-slate-500">Welcome back, <span className="font-semibold text-blue-600">{currentUser.displayName}</span></p><p className="mt-1 text-xs font-medium text-slate-400">Branch IT Control workspace</p></div>
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white px-3.5 py-2.5 text-slate-600 shadow-sm"><AppIcon name="calendar" className="h-4 w-4 text-blue-600" /><div><p className="text-[10px] font-semibold uppercase tracking-wide text-slate-400">Today</p><p className="text-xs font-semibold">{new Date().toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short", year: "numeric" })}</p></div></div>
            <button onClick={() => syncFromServer(true)} className="flex h-11 w-11 items-center justify-center rounded-xl border border-slate-200 bg-white text-slate-600 shadow-sm hover:border-blue-200 hover:text-blue-600" title="Sync shared data"><AppIcon name="sync" className="h-5 w-5" /></button>
            <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white px-3 py-2 shadow-sm"><div className="flex h-9 w-9 items-center justify-center rounded-full bg-indigo-700 text-sm font-bold text-white">{String(currentUser.displayName || "U").slice(0,1).toUpperCase()}</div><div className="pr-2"><p className="text-xs font-bold text-slate-900">{currentUser.displayName}</p><p className="text-[10px] text-slate-500">{isAdmin ? "Super Admin" : "User"}</p></div></div>
          </div>
        </header>
        <div className="mobile-top-card mb-3 rounded-2xl bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 p-4 text-white shadow-lg lg:hidden"><div className="flex items-center justify-between gap-3"><div><p className="text-[10px] font-semibold uppercase tracking-[0.16em] text-blue-200">Branch IT Control</p><h1 className="mt-1 text-base font-bold">Welcome, {currentUser.displayName}</h1><p className="mt-0.5 text-xs text-slate-300">{isAdmin ? "Super Admin" : "View only"}</p></div><div className="flex gap-2"><button onClick={() => syncFromServer(true)} className="rounded-xl bg-white/10 p-2.5"><AppIcon name="sync" className="h-4 w-4" /></button><button onClick={logout} className="rounded-xl bg-white p-2.5 text-slate-900"><AppIcon name="logout" className="h-4 w-4" /></button></div></div></div>
        <div className="px-0 pb-6 lg:px-5 lg:pt-4">
          
          <div className="mobile-menu mt-3 lg:hidden"><select value={tab} onChange={(e) => setTab(e.target.value)} className="w-full rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 outline-none">{menuItems.map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></div>

        {tab === "dashboard" && <DashboardSection completedProjects={completedProjects} ongoingProjects={ongoingProjects} upcomingProjects={upcomingProjects} branches={branches} assets={assets} openBranch={openBranch} totals={{ locations: locations.length, branches: branches.length, assets: assets.reduce((sum, a) => sum + Math.max(1, Number(a.quantity || 1)), 0), pending: pendingHandover }} setTab={setTab} openModal={openModal} exportDatabase={exportDatabase} isAdmin={isAdmin} />}
        {tab === "contribution" && <UserPortfolio profiles={profiles} selectedContributor={selectedContributor} setSelectedContributor={setSelectedContributor} selectedProfile={selectedProfile} isAdmin={isAdmin} openProfile={(p) => openModal("profile", p)} projects={contributorProjects} faults={contributorFaults} handovers={contributorHandovers} pullouts={contributorPullouts} reviews={contributorCctvReviews} branches={branches} openBranch={openBranch} />}
        {tab === "groups" && <GroupsSection groups={groups} branches={branches} assets={assets} isAdmin={isAdmin} openModal={openModal} />}
        {tab === "locations" && <LocationsSection locations={locations} branches={branches} isAdmin={isAdmin} search={search.location} setSearch={(v) => setSearch((x) => ({ ...x, location: v }))} openAdd={() => openModal("location")} openEdit={(location) => openModal("location", location)} openImport={() => openModal("locationImport")} openTemplate={() => openModal("template", { type: "location" })} deleteLocation={deleteLocation} />}
        {tab === "branches" && <BranchesSection branches={filteredBranches} assets={assets} search={search.branch} setSearch={(v) => setSearch((x) => ({ ...x, branch: v }))} openBranch={openBranch} isAdmin={isAdmin} openTemplate={() => openModal("template", { type: "branch" })} openImport={() => openModal("branchImport")} />}
        {tab === "branch-detail" && <BranchDetail selectedBranch={selectedBranch} assets={filteredSelectedBranchAssets} allAssets={selectedBranchAssets} network={selectedNetwork} handover={selectedHandover} pullout={selectedPullout} faults={selectedBranchFaults} isAdmin={isAdmin} subTab={branchSubTab} setSubTab={setBranchSubTab} openModal={openModal} deleteBranch={deleteBranch} deleteAsset={deleteAsset} deleteAssetsBulk={deleteAssetsBulk} deleteFault={deleteFault} printReport={printReport} updateBranchLogo={updateBranchLogo} removeBranchLogo={removeBranchLogo} groupBranding={selectedGroupBranding} branchAssetSearch={search.branchAsset} setBranchAssetSearch={(v) => setSearch((x) => ({ ...x, branchAsset: v }))} assetTemplates={assetTemplates} />}
        {tab === "assets" && <AssetsSection assets={filteredAssets} branches={branches} search={search.asset} setSearch={(v) => setSearch((x) => ({ ...x, asset: v }))} openBranch={openBranch} isAdmin={isAdmin} assetProducts={assetProducts} assetTemplates={assetTemplates} allAssets={assets} assetCategories={assetCategories} openModal={openModal} deleteProduct={deleteProduct} deleteProductsBulk={deleteProductsBulk} setProductsActive={setProductsActive} deleteAssetTemplate={deleteAssetTemplate} />}
        {tab === "cctv" && <CctvSection records={nvrRecords} assets={assets} isAdmin={isAdmin} search={search.cctv} setSearch={(v) => setSearch((x) => ({ ...x, cctv: v }))} openBranch={openBranch} branches={branches} openCctv={(n) => openModal("cctv", n)} />}
        {tab === "vault" && <CredentialVaultSection credentials={credentials} branches={branches} assets={assets} isAdmin={isAdmin} vaultSettings={vaultSettings} vaultUnlocked={vaultUnlocked} initializeVault={initializeVault} unlockVault={unlockVault} lockVault={lockVault} revealSecret={revealCredentialSecret} openCredential={(item) => openModal("credential", item)} deleteCredential={deleteCredential} downloadTemplate={downloadCredentialTemplate} exportRegister={exportCredentialRegister} importRef={credentialImportRef} importExcel={importCredentialExcel} />}
        {tab === "faults" && <FaultsSection faults={filteredFaults} isAdmin={isAdmin} search={search.fault} setSearch={(v) => setSearch((x) => ({ ...x, fault: v }))} faultView={faultView} setFaultView={setFaultView} range={reportRange} setRange={setReportRange} openModal={openModal} openBranch={openBranch} deleteFault={deleteFault} downloadCSV={downloadFaultCSV} />}
        {tab === "handover" && <section className="mt-6 grid gap-4 xl:grid-cols-2"><div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><h3 className="text-lg font-bold">Global Reports</h3><p className="mt-1 text-sm text-slate-500">Branch-specific reports are inside each branch record.</p>{["All Branch Asset Register", "All Branch Network Summary", "Warranty Expiry Report", "CCTV Asset Register"].map((x) => <button key={x} onClick={exportDatabase} className="mt-3 flex w-full items-center justify-between rounded-2xl border border-slate-200 bg-white p-4 text-left text-sm font-semibold hover:bg-slate-50"><span>{x}</span><span>Download</span></button>)}</div></section>}
        {tab === "settings" && <SettingsSection theme={theme} setTheme={setTheme} isAdmin={isAdmin} users={users} forms={forms} updateForm={updateForm} setUsers={setUsers} currentUser={currentUser} setCurrentUser={setCurrentUser} deleteUser={deleteUser} exportDatabase={exportDatabase} restoreInputRef={restoreInputRef} restoreDatabaseFile={restoreDatabaseFile} resetData={resetData} syncFromServer={syncFromServer} serverMessage={serverMessage} cleanImportedExcelRemarks={cleanImportedExcelRemarks} assetProducts={assetProducts} assetCategories={assetCategories} setAssetCategories={setAssetCategories} assets={assets} openModal={openModal} deleteProduct={deleteProduct} resetRequests={resetRequests} setResetRequests={setResetRequests} />}
        </div>
      </main>
    </div>
    {confirmDialog && <ConfirmDialog dialog={confirmDialog} onCancel={() => setConfirmDialog(null)} onConfirm={() => { const action = confirmDialog.onConfirm; setConfirmDialog(null); if (typeof action === "function") action(); }} />}
    {modal && <ModalRouter modal={modal} close={closeModal} forms={forms} updateForm={updateForm} formError={formError} submit={{ location: submitLocation, branch: submitBranch, asset: submitAsset, product: submitProduct, assetTemplate: submitAssetTemplate, network: submitNetwork, fault: submitFault, handover: submitHandover, pullout: submitPullout, profile: submitProfile, groupBranding: submitGroupBranding, credential: submitCredential, cctv: submitCctv }} groups={groups} selectedBranch={selectedBranch} selectedBranchAssets={selectedBranchAssets} locations={locations} branches={branches} selectedNetwork={selectedNetwork} editId={editId} setModal={setModal} openModal={openModal} excelImport={excelImport} excelActions={{ handleExcelAssetFile, loadExcelSheet, updateExcelMapping, updateExcelDuplicateMode, buildExcelImportPreview, confirmExcelAssetImport }} branchImport={branchImport} branchActions={{ handleBranchExcelFile, loadBranchExcelSheet, updateBranchImportMapping, updateBranchImportDuplicateMode, buildBranchImportPreview, confirmBranchImport }} locationImport={locationImport} locationActions={{ handleLocationExcelFile, loadLocationExcelSheet, updateLocationImportMapping, updateLocationImportDuplicateMode, buildLocationImportPreview, confirmLocationImport }} productImport={productImport} productActions={{ handleProductExcelFile, loadProductExcelSheet, updateProductImportMapping, updateProductImportDuplicateMode, buildProductImportPreview, confirmProductImport }} templateBuilder={templateBuilder} templateActions={{ updateTemplateField, selectTemplatePreset, downloadSelectedTemplate }} assetProducts={assetProducts} assetCategories={assetCategories} addAssetCategory={addAssetCategory} assetTemplates={assetTemplates} applyBranchAssetTemplate={applyBranchAssetTemplate} vaultUnlocked={vaultUnlocked} allAssets={assets} credentials={credentials} />}
  </div>;
}

function CredentialVaultSection({ credentials, branches, assets, isAdmin, vaultSettings, vaultUnlocked, initializeVault, unlockVault, lockVault, revealSecret, openCredential, deleteCredential, downloadTemplate, exportRegister, importRef, importExcel }) {
  const [activeType, setActiveType] = useState("shared");
  const [searchText, setSearchText] = useState("");
  const [scopeFilter, setScopeFilter] = useState("All");
  const [passphrase, setPassphrase] = useState("");
  const [confirmation, setConfirmation] = useState("");
  const [vaultError, setVaultError] = useState("");
  const [busy, setBusy] = useState(false);
  const [revealed, setRevealed] = useState({});
  const records = credentials.map(normalizeCredentialRecord);

  if (!isAdmin) {
    return <section className="mt-6 rounded-3xl border border-amber-200 bg-amber-50 p-8 text-center shadow-sm"><div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-amber-100 text-amber-700"><AppIcon name="lock" className="h-7 w-7" /></div><h3 className="mt-4 text-xl font-bold text-amber-900">Credential Vault Restricted</h3><p className="mx-auto mt-2 max-w-xl text-sm leading-6 text-amber-800">Only administrator accounts can access shared accounts, device passwords and access-register records.</p></section>;
  }

  async function handleSetup(e) {
    e.preventDefault(); setBusy(true); setVaultError("");
    const result = await initializeVault(passphrase, confirmation);
    if (!result.ok) setVaultError(result.message); else { setPassphrase(""); setConfirmation(""); }
    setBusy(false);
  }
  async function handleUnlock(e) {
    e.preventDefault(); setBusy(true); setVaultError("");
    const result = await unlockVault(passphrase);
    if (!result.ok) setVaultError(result.message); else setPassphrase("");
    setBusy(false);
  }

  if (!vaultSettings) {
    return <section className="mt-6 grid gap-5 xl:grid-cols-[1.15fr_0.85fr]"><div className="rounded-3xl border border-slate-200 bg-white p-7 shadow-sm"><div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-blue-50 text-blue-700"><AppIcon name="vault" className="h-7 w-7" /></div><h3 className="mt-5 text-2xl font-bold">Create Credential Vault</h3><p className="mt-2 max-w-2xl text-sm leading-6 text-slate-500">Create one master passphrase. Passwords are encrypted in your browser before they are saved. The passphrase itself is never stored.</p><form onSubmit={handleSetup} className="mt-6 max-w-xl space-y-4"><Field label="New Vault Passphrase" required><PasswordInput value={passphrase} onChange={setPassphrase} placeholder="Minimum 8 characters" /></Field><Field label="Confirm Passphrase" required><PasswordInput value={confirmation} onChange={setConfirmation} placeholder="Enter the same passphrase" /></Field>{vaultError && <ErrorBox text={vaultError} />}<button disabled={busy} className="rounded-2xl bg-slate-900 px-5 py-3 text-sm font-semibold text-white disabled:opacity-60">{busy ? "Creating Vault..." : "Create and Unlock Vault"}</button></form></div><div className="rounded-3xl border border-amber-200 bg-amber-50 p-7"><h4 className="text-lg font-bold text-amber-900">Keep the passphrase safely</h4><p className="mt-3 text-sm leading-6 text-amber-800">There is no recovery option for the encryption passphrase. Losing it means encrypted passwords cannot be opened.</p><div className="mt-5 rounded-2xl border border-amber-200 bg-white/70 p-4 text-sm text-amber-900"><b>Recommended:</b> store it in the approved company password manager and limit it to authorized IT management.</div></div></section>;
  }

  if (!vaultUnlocked) {
    return <section className="mt-6 flex min-h-[520px] items-center justify-center"><div className="w-full max-w-lg rounded-3xl border border-slate-200 bg-white p-8 text-center shadow-sm"><div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-slate-900 text-white"><AppIcon name="lock" className="h-8 w-8" /></div><h3 className="mt-5 text-2xl font-bold">Unlock Credential Vault</h3><p className="mt-2 text-sm leading-6 text-slate-500">Enter the vault passphrase to reveal, add, edit or import encrypted credentials.</p><form onSubmit={handleUnlock} className="mt-6 space-y-4 text-left"><Field label="Vault Passphrase" required><PasswordInput value={passphrase} onChange={setPassphrase} placeholder="Enter vault passphrase" /></Field>{vaultError && <ErrorBox text={vaultError} />}<button disabled={busy} className="w-full rounded-2xl bg-slate-900 px-5 py-3 text-sm font-semibold text-white disabled:opacity-60">{busy ? "Unlocking..." : "Unlock Vault"}</button></form></div></section>;
  }

  const sharedAccounts = records.filter((item) => item.type === "shared");
  const linkedSharedAccount = (item) => sharedAccounts.find((account) => Number(account.id) === Number(item.linkedCredentialId));
  const branchFor = (item) => branches.find((branch) => branch.id === Number(item.branchId));
  const assetFor = (item) => assets.find((asset) => asset.id === Number(item.linkedAssetId));
  const recordScope = (item) => {
    if (item.type === "shared") return item.scopeType || "Other";
    if (item.type === "access") return linkedSharedAccount(item)?.scopeType || item.departmentOperation || "Other";
    const branch = branchFor(item);
    const text = `${branch?.branchType || ""} ${branch?.groupName || ""} ${branch?.conceptName || ""} ${branch?.branchName || ""}`.toLowerCase();
    if (text.includes("din tai fung") || text.includes("dtf")) return "DTF";
    if (text.includes("stockroom")) return "Stockroom";
    if (text.includes("office") || text.includes("head office") || text.includes("ho -")) return "Office";
    if (text.includes("warehouse")) return "Warehouse";
    if (text.includes("qsr")) return "QSR";
    return branch?.branchType || "Other";
  };
  const filtered = records.filter((item) => {
    if (item.type !== activeType) return false;
    if (scopeFilter !== "All" && recordScope(item) !== scopeFilter) return false;
    const branch = branchFor(item);
    const linkedAccount = linkedSharedAccount(item);
    const linkedAsset = assetFor(item);
    const haystack = `${item.category} ${item.deviceName} ${item.platform} ${item.accountName} ${item.scopeType} ${item.scopeName} ${item.ipAddress} ${item.url} ${item.username} ${item.email} ${item.owner} ${item.jobRole} ${item.departmentOperation} ${item.permissionLevel} ${item.status} ${item.notes} ${branch?.branchName || ""} ${linkedAccount?.accountName || ""} ${linkedAsset?.serialNumber || ""}`.toLowerCase();
    return haystack.includes(searchText.toLowerCase());
  });

  const todayValue = today();
  const counts = {
    shared: sharedAccounts.length,
    device: records.filter((item) => item.type === "device").length,
    access: records.filter((item) => item.type === "access").length,
    activeAccess: records.filter((item) => item.type === "access" && item.status === "Active").length,
    reviewDue: records.filter((item) => item.type !== "access" && item.status === "Active" && item.nextPasswordChange && item.nextPasswordChange <= todayValue).length,
    inactive: records.filter((item) => item.type !== "access" && ["Inactive", "Revoked", "Expired"].includes(item.status)).length,
    revoked: records.filter((item) => item.type === "access" && item.status === "Revoked").length,
  };

  async function reveal(item) {
    try {
      const secret = await revealSecret(item);
      setRevealed((current) => ({ ...current, [item.id]: secret || "No password saved" }));
      window.setTimeout(() => setRevealed((current) => { const next = { ...current }; delete next[item.id]; return next; }), 30000);
    } catch (error) { alert(error.message); }
  }
  async function copySecret(item) {
    try {
      const secret = revealed[item.id] || await revealSecret(item);
      if (!secret) return alert("No password is stored for this record.");
      await navigator.clipboard.writeText(secret);
      alert("Password copied to clipboard.");
    } catch (error) { alert(`Unable to copy password.\n\n${error.message}`); }
  }
  const branchName = (item) => branchFor(item)?.branchName || "Not linked";
  const passwordCell = (item) => item.secretCipher ? <div className="min-w-[150px]"><p className="break-all font-mono text-xs">{revealed[item.id] || "••••••••••"}</p><div className="mt-1.5 flex gap-1.5"><button onClick={() => reveal(item)} className="rounded-lg border border-slate-300 px-2 py-1 text-[10px] font-semibold"><AppIcon name={revealed[item.id] ? "eyeOff" : "eye"} className="mr-1 inline h-3 w-3" />{revealed[item.id] ? "Visible" : "Reveal"}</button><button onClick={() => copySecret(item)} className="rounded-lg border border-slate-300 px-2 py-1 text-[10px] font-semibold"><AppIcon name="copy" className="mr-1 inline h-3 w-3" />Copy</button></div></div> : <span className="text-xs text-slate-400">Not saved</span>;
  const actions = (item) => <div className="flex flex-wrap gap-1.5"><button onClick={() => openCredential(item)} className="rounded-lg border border-slate-300 px-2.5 py-1.5 text-xs font-semibold">Edit</button><button onClick={() => deleteCredential(item)} className="rounded-lg border border-red-200 bg-red-50 px-2.5 py-1.5 text-xs font-semibold text-red-700">Delete</button></div>;

  return <section className="mt-6 space-y-4">
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between"><div className="flex items-start gap-3"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-blue-50 text-blue-700"><AppIcon name="vault" className="h-5 w-5" /></div><div><h3 className="text-xl font-bold">Credential Vault</h3><p className="mt-1 text-sm text-slate-500">Shared accounts, device credentials and user-access assignments in one encrypted register.</p></div></div><div className="flex flex-wrap gap-2"><button onClick={() => openCredential(null)} className="rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white">+ Add Record</button><button onClick={() => importRef.current?.click()} className="rounded-2xl border border-emerald-300 bg-emerald-50 px-4 py-2.5 text-sm font-semibold text-emerald-700">Import Excel</button><input ref={importRef} type="file" accept=".xlsx,.xls" onChange={(e) => importExcel(e.target.files?.[0])} className="hidden" /><button onClick={downloadTemplate} className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold">Download Template</button><button onClick={exportRegister} className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold">Export Register</button><button onClick={lockVault} className="rounded-2xl border border-amber-300 bg-amber-50 px-4 py-2.5 text-sm font-semibold text-amber-800">Lock Vault</button></div></div></div>
    <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-6"><Card title="Shared Accounts" value={counts.shared} note="QSR, DTF, office and group accounts" icon="groups" accent="purple" /><Card title="Device Credentials" value={counts.device} note="NVR, server and network devices" icon="assets" accent="blue" /><Card title="Active User Access" value={counts.activeAccess} note="People with current access" icon="user" accent="green" /><Card title="Password Review Due" value={counts.reviewDue} note="Active accounts due for review" icon="calendar" accent="amber" /><Card title="Inactive Accounts" value={counts.inactive} note="Inactive, expired or revoked" icon="lock" accent="rose" /><Card title="Revoked Access" value={counts.revoked} note="User access already removed" icon="faults" accent="teal" /></div>
    <div className="rounded-3xl border border-slate-200 bg-white p-4 shadow-sm"><div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between"><div className="flex flex-wrap gap-2">{Object.entries(credentialTypeLabels).map(([key, label]) => <button key={key} onClick={() => setActiveType(key)} className={`rounded-2xl px-4 py-2 text-sm font-semibold ${activeType === key ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-700"}`}>{label} <span className="ml-1 opacity-70">{counts[key]}</span></button>)}</div><div className="flex w-full flex-col gap-2 sm:flex-row xl:w-auto"><select value={scopeFilter} onChange={(e) => setScopeFilter(e.target.value)} className="rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm outline-none"><option>All</option>{credentialScopes.map((scope) => <option key={scope}>{scope}</option>)}</select><input value={searchText} onChange={(e) => setSearchText(e.target.value)} placeholder="Search vault records..." className="w-full rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm outline-none sm:min-w-[280px]" /></div></div></div>

    {activeType === "shared" && <Table headers={["Platform / Scope", "Account", "Login", "Password", "Owner / MFA", "Review", "Status", "Actions"]} rows={filtered} renderDesktop={(item) => <tr key={item.id} className="hover:bg-slate-50"><td className="px-4 py-3"><p className="font-semibold">{item.platform || "Shared Account"}</p><p className="text-xs text-slate-500">{item.scopeType || "Other"}{item.scopeName ? ` • ${item.scopeName}` : ""}</p></td><td className="px-4 py-3 font-semibold">{item.accountName || "-"}</td><td className="px-4 py-3">{item.email || item.username || "-"}<p className="text-xs text-slate-500">{item.email && item.username ? item.username : ""}</p></td><td className="px-4 py-3">{passwordCell(item)}</td><td className="px-4 py-3">{item.owner || "IT Department"}<p className="text-xs text-slate-500">MFA: {item.mfaStatus || "Not Enabled"}</p></td><td className="px-4 py-3">{item.nextPasswordChange || "Not set"}<p className="text-xs text-slate-500">Last: {item.lastPasswordChange || "-"}</p></td><td className="px-4 py-3"><Badge type={item.status === "Active" ? "success" : "warning"}>{item.status}</Badge></td><td className="px-4 py-3">{actions(item)}</td></tr>} renderMobile={(item) => <MobileCard key={item.id} kicker={`${item.platform || "Shared Account"} • ${item.scopeType || "Other"}`} title={item.accountName || item.username || item.email} subtitle={item.scopeName || item.owner} status={<Badge type={item.status === "Active" ? "success" : "warning"}>{item.status}</Badge>} details={[["Login", item.email || item.username || "-"], ["MFA", item.mfaStatus || "Not Enabled"], ["Password", passwordCell(item)], ["Next Review", item.nextPasswordChange || "Not set"]]} actions={actions(item)} />} />}

    {activeType === "device" && <Table headers={["Branch / Device", "Linked Asset", "Address", "Username", "Password", "Review", "Status", "Actions"]} rows={filtered} renderDesktop={(item) => { const linkedAsset = assetFor(item); return <tr key={item.id} className="hover:bg-slate-50"><td className="px-4 py-3"><p className="font-semibold">{branchName(item)}</p><p className="text-xs text-slate-500">{item.deviceName || item.category}</p></td><td className="px-4 py-3">{linkedAsset?.model || item.category || "-"}<p className="text-xs text-slate-500">{linkedAsset?.serialNumber || "No asset linked"}</p></td><td className="px-4 py-3 text-slate-600">{item.ipAddress || item.url || "-"}<p className="text-xs text-slate-400">{item.port ? `Port ${item.port}` : ""}</p></td><td className="px-4 py-3">{item.username || item.email || "-"}</td><td className="px-4 py-3">{passwordCell(item)}</td><td className="px-4 py-3">{item.nextPasswordChange || "Not set"}<p className="text-xs text-slate-500">Last: {item.lastPasswordChange || "-"}</p></td><td className="px-4 py-3"><Badge type={item.status === "Active" ? "success" : "warning"}>{item.status}</Badge></td><td className="px-4 py-3">{actions(item)}</td></tr>; }} renderMobile={(item) => { const linkedAsset = assetFor(item); return <MobileCard key={item.id} kicker={branchName(item)} title={item.deviceName || item.category} subtitle={item.ipAddress || item.url} status={<Badge type={item.status === "Active" ? "success" : "warning"}>{item.status}</Badge>} details={[["Asset", linkedAsset ? `${linkedAsset.model || linkedAsset.deviceName} • ${linkedAsset.serialNumber || "No serial"}` : "Not linked"], ["Username", item.username || item.email || "-"], ["Password", passwordCell(item)], ["Next Review", item.nextPasswordChange || "Not set"]]} actions={actions(item)} />; }} />}

    {activeType === "access" && <Table headers={["Person", "Linked Shared Account", "Role / Operation", "Permission", "Granted / Approved", "Status", "Actions"]} rows={filtered} renderDesktop={(item) => { const linked = linkedSharedAccount(item); return <tr key={item.id} className="hover:bg-slate-50"><td className="px-4 py-3"><p className="font-semibold">{item.owner || "-"}</p><p className="text-xs text-slate-500">{item.email || item.mobile || ""}</p></td><td className="px-4 py-3"><p className="font-semibold">{linked?.accountName || "Unlinked account"}</p><p className="text-xs text-slate-500">{linked?.platform || "-"} • {linked?.scopeType || item.departmentOperation || "-"}</p></td><td className="px-4 py-3">{item.jobRole || "-"}<p className="text-xs text-slate-500">{item.departmentOperation || "-"}</p></td><td className="px-4 py-3">{item.permissionLevel || "Viewer"}</td><td className="px-4 py-3">{item.grantedDate || "-"}<p className="text-xs text-slate-500">{item.approvedBy ? `By ${item.approvedBy}` : ""}</p></td><td className="px-4 py-3"><Badge type={item.status === "Active" ? "success" : item.status === "Revoked" ? "danger" : "warning"}>{item.status}</Badge></td><td className="px-4 py-3">{actions(item)}</td></tr>; }} renderMobile={(item) => { const linked = linkedSharedAccount(item); return <MobileCard key={item.id} kicker={linked?.platform || "Access Register"} title={item.owner || "-"} subtitle={linked?.accountName || "Unlinked account"} status={<Badge type={item.status === "Active" ? "success" : item.status === "Revoked" ? "danger" : "warning"}>{item.status}</Badge>} details={[["Role", item.jobRole || "-"], ["Operation", item.departmentOperation || linked?.scopeType || "-"], ["Permission", item.permissionLevel || "Viewer"], ["Granted", item.grantedDate || "-"]]} actions={actions(item)} />; }} />}
  </section>;
}

function DashboardSection({ completedProjects, ongoingProjects, upcomingProjects, branches, assets, openBranch, totals, setTab, openModal, exportDatabase, isAdmin }) {
  const categoryTotals = useMemo(() => {
    const map = {};
    assets.forEach((asset) => {
      const key = asset.category || "Other";
      map[key] = (map[key] || 0) + Math.max(1, Number(asset.quantity || 1));
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [assets]);
  const recentBranches = [...branches].slice(-4).reverse();
  const recentAssets = [...assets].slice(-4).reverse();
  const pendingBranches = branches.filter((branch) => Number(branch.handover || 0) < 100).slice(0, 4);
  const completedCount = branches.filter((branch) => branch.status === "Live" || Number(branch.handover || 0) === 100).length;
  const ongoingCount = branches.filter((branch) => ["Installation Started", "Pending Handover", "Opening Today"].includes(branch.status) && Number(branch.handover || 0) < 100).length;
  const upcomingCount = branches.filter((branch) => branch.status === "Planning").length;
  const otherCount = Math.max(0, branches.length - completedCount - ongoingCount - upcomingCount);
  const percentage = (value, total) => total ? Math.round((value / total) * 1000) / 10 : 0;
  const branchStatusRows = [
    ["Completed", completedCount, "#22c55e"],
    ["Ongoing", ongoingCount, "#f59e0b"],
    ["Upcoming", upcomingCount, "#6366f1"],
    ["Other", otherCount, "#94a3b8"],
  ];
  const branchSegments = branchStatusRows.map(([label, value, color]) => ({ label, value, color }));
  const assetColors = ["#2563eb", "#22c55e", "#f97316", "#8b5cf6", "#06b6d4"];
  const primaryCategories = categoryTotals.slice(0, 4);
  const remainingAssets = categoryTotals.slice(4).reduce((sum, [, value]) => sum + value, 0);
  const visibleCategories = remainingAssets > 0 ? [...primaryCategories, ["Other", remainingAssets]] : primaryCategories;
  const assetSegments = visibleCategories.map(([label, value], index) => ({ label, value, color: assetColors[index] }));
  const Kpis = [
    ["Completed Projects", completedProjects.length, "Live or fully handed over", "check", "blue", () => setTab("branches")],
    ["Ongoing Projects", ongoingProjects.length, "Installation and handover", "briefcase", "amber", () => setTab("branches")],
    ["Upcoming Projects", upcomingProjects.length, "Branches in planning", "calendar", "purple", () => setTab("branches")],
    ["Locations", totals.locations, "Mall and site master", "location", "green", () => setTab("locations")],
    ["Branches", totals.branches, "Total branch records", "branches", "blue", () => setTab("branches")],
    ["IT Assets", totals.assets, "Asset quantity across branches", "assets", "teal", () => setTab("assets")],
  ];
  const branchForAsset = (asset) => branches.find((branch) => branch.id === asset.branchId);
  return <section className="mt-4 space-y-4">
    <div><h2 className="text-2xl font-bold tracking-tight text-slate-950">Dashboard</h2><p className="mt-1 text-sm text-slate-500">Overview of branch projects, locations and IT asset activity.</p></div>
    <div className="dashboard-mobile-grid grid gap-3 md:grid-cols-2 xl:grid-cols-6">{Kpis.map(([title, value, note, icon, accent, action]) => <Card key={title} title={title} value={value} note={note} icon={icon} accent={accent} onClick={action} />)}</div>
    <div className="grid gap-4 xl:grid-cols-[1fr_1.05fr_0.95fr]">
      <div className="professional-page-card rounded-2xl p-5">
        <div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-xl bg-blue-50 text-blue-600"><AppIcon name="branches" className="h-4 w-4" /></div><div><h3 className="font-bold text-slate-950">Branch Overview</h3><p className="mt-0.5 text-xs text-slate-500">Accurate project status distribution</p></div></div>
        <div className="mt-5 grid items-center gap-5 sm:grid-cols-[132px_1fr]"><DonutChart segments={branchSegments} total={branches.length} label="Total Branches" /><div className="space-y-2.5">{branchStatusRows.map(([label, count, color]) => <div key={label} className="flex items-center justify-between gap-3 text-sm"><span className="flex min-w-0 items-center gap-2 text-slate-600"><span className="h-2.5 w-2.5 shrink-0 rounded-full" style={{ backgroundColor: color }} /><span className="truncate">{label}</span></span><span className="shrink-0 font-semibold text-slate-900">{count} <span className="font-normal text-slate-400">({percentage(count, branches.length)}%)</span></span></div>)}</div></div>
        <button onClick={() => setTab("branches")} className="mt-5 text-xs font-semibold text-blue-600 hover:text-blue-700">View all branches →</button>
      </div>
      <div className="professional-page-card rounded-2xl p-5">
        <div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-xl bg-violet-50 text-violet-600"><AppIcon name="assets" className="h-4 w-4" /></div><div><h3 className="font-bold text-slate-950">IT Assets by Category</h3><p className="mt-0.5 text-xs text-slate-500">Top categories with remaining assets grouped as Other</p></div></div>
        <div className="mt-5 grid items-center gap-5 sm:grid-cols-[132px_1fr]"><DonutChart segments={assetSegments} total={totals.assets} label="Total Assets" /><div className="space-y-2.5">{visibleCategories.length === 0 ? <p className="text-sm text-slate-500">No asset data available.</p> : visibleCategories.map(([name, count], index) => <div key={name} className="flex items-center justify-between gap-3 text-sm"><span className="flex min-w-0 items-center gap-2 text-slate-600"><span className="h-2.5 w-2.5 shrink-0 rounded-full" style={{ backgroundColor: assetColors[index] }} /><span className="truncate">{name}</span></span><span className="shrink-0 font-semibold text-slate-900">{count} <span className="font-normal text-slate-400">({percentage(count, totals.assets)}%)</span></span></div>)}</div></div>
        <button onClick={() => setTab("assets")} className="mt-5 text-xs font-semibold text-blue-600 hover:text-blue-700">View all assets →</button>
      </div>
      <div className="overflow-hidden rounded-2xl border border-blue-100 bg-gradient-to-br from-blue-50 via-white to-indigo-50 p-5 shadow-[0_10px_30px_rgba(37,99,235,0.08)]"><div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-xl bg-blue-600 text-white"><AppIcon name="plus" className="h-4 w-4" /></div><div><h3 className="font-bold text-slate-950">Quick Actions</h3><p className="mt-0.5 text-xs text-slate-500">Common operational tasks</p></div></div><div className="mt-5 grid gap-3">{isAdmin && <><button onClick={() => openModal("branch")} className="flex items-center justify-between rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm hover:border-blue-200"><span className="flex items-center gap-3"><span className="rounded-lg bg-blue-50 p-2 text-blue-600"><AppIcon name="plus" className="h-4 w-4" /></span>Add New Branch</span><span>›</span></button><button onClick={() => openModal("branchImport")} className="flex items-center justify-between rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm hover:border-blue-200"><span className="flex items-center gap-3"><span className="rounded-lg bg-emerald-50 p-2 text-emerald-600"><AppIcon name="upload" className="h-4 w-4" /></span>Upload Branches Excel</span><span>›</span></button><button onClick={() => openModal("asset")} className="flex items-center justify-between rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm hover:border-blue-200"><span className="flex items-center gap-3"><span className="rounded-lg bg-violet-50 p-2 text-violet-600"><AppIcon name="assets" className="h-4 w-4" /></span>Add IT Asset</span><span>›</span></button></>}<button onClick={exportDatabase} className="flex items-center justify-between rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm hover:border-blue-200"><span className="flex items-center gap-3"><span className="rounded-lg bg-amber-50 p-2 text-amber-600"><AppIcon name="download" className="h-4 w-4" /></span>Download Report</span><span>›</span></button></div></div>
    </div>
    <div className="grid gap-4 xl:grid-cols-3">
      <div className="professional-page-card rounded-2xl p-5"><div className="flex items-center justify-between"><div><h3 className="font-bold text-slate-950">Recent Branches</h3><p className="text-xs text-slate-500">Latest branch records</p></div><button onClick={() => setTab("branches")} className="text-xs font-semibold text-blue-600">View all</button></div><div className="mt-4 divide-y divide-slate-100">{recentBranches.map((branch) => <button key={branch.id} onClick={() => openBranch(branch)} className="flex w-full items-center justify-between gap-3 py-3 text-left"><div className="min-w-0"><p className="truncate text-sm font-semibold text-slate-900">{branch.branchName}</p><p className="mt-0.5 truncate text-xs text-slate-500">{branch.locationName}, {branch.city}</p></div><Badge type={branch.status === "Live" ? "success" : branch.status === "Planning" ? "info" : "warning"}>{branch.status}</Badge></button>)}</div></div>
      <div className="professional-page-card rounded-2xl p-5"><div className="flex items-center justify-between"><div><h3 className="font-bold text-slate-950">Recent Assets Added</h3><p className="text-xs text-slate-500">Latest asset records</p></div><button onClick={() => setTab("assets")} className="text-xs font-semibold text-blue-600">View all</button></div><div className="mt-4 divide-y divide-slate-100">{recentAssets.map((asset) => { const branch = branchForAsset(asset); return <div key={asset.id} className="flex items-center gap-3 py-3"><span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-600"><AppIcon name="assets" className="h-4 w-4" /></span><div className="min-w-0"><p className="truncate text-sm font-semibold text-slate-900">{asset.deviceName || asset.category}</p><p className="mt-0.5 truncate text-xs text-slate-500">{branch?.branchName || "Unassigned"} • {asset.brand || "Brand not set"}</p></div></div>; })}</div></div>
      <div className="professional-page-card rounded-2xl p-5"><div className="flex items-center justify-between"><div><h3 className="font-bold text-slate-950">Pending Handovers</h3><p className="text-xs text-slate-500">Branches below 100%</p></div><Badge type="warning">{totals.pending}</Badge></div><div className="mt-4 divide-y divide-slate-100">{pendingBranches.length === 0 ? <p className="py-6 text-center text-sm text-slate-500">No pending handovers.</p> : pendingBranches.map((branch) => <button key={branch.id} onClick={() => openBranch(branch)} className="flex w-full items-center justify-between gap-3 py-3 text-left"><div className="min-w-0"><p className="truncate text-sm font-semibold text-slate-900">{branch.branchName}</p><p className="mt-0.5 truncate text-xs text-slate-500">{branch.locationName}, {branch.city}</p></div><span className="rounded-full bg-amber-100 px-2.5 py-1 text-xs font-semibold text-amber-700">{Number(branch.handover || 0)}%</span></button>)}</div></div>
    </div>
  </section>;
}

function UserPortfolio({ profiles, selectedContributor, setSelectedContributor, selectedProfile, isAdmin, openProfile, projects = [], faults = [], handovers = [], pullouts = [], reviews = [], branches, openBranch }) {
  const [period, setPeriod] = useState("all");
  const now = new Date();
  const withinPeriod = (value) => {
    if (period === "all" || !value) return true;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return false;
    if (period === "month") return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
    if (period === "quarter") return date.getFullYear() === now.getFullYear() && Math.floor(date.getMonth() / 3) === Math.floor(now.getMonth() / 3);
    if (period === "year") return date.getFullYear() === now.getFullYear();
    return true;
  };
  const filteredProjects = projects.filter((item) => withinPeriod(item.completedDate || item.openingDate));
  const filteredFaults = faults.filter((item) => withinPeriod(item.faultDate));
  const filteredHandovers = handovers.filter((item) => withinPeriod(item.handoverDate));
  const filteredPullouts = pullouts.filter((item) => withinPeriod(item.pulloutDate));
  const filteredReviews = reviews.filter((item) => withinPeriod(item.lastRecordingReviewDate));
  const completed = filteredProjects.filter((item) => item.status === "Live" || Number(item.handover || 0) === 100);
  const active = filteredProjects.filter((item) => !["Live", "Closed"].includes(item.status) && Number(item.handover || 0) < 100);
  const closedFaults = filteredFaults.filter((item) => item.status === "Closed");
  const openFaults = filteredFaults.filter((item) => item.status !== "Closed");
  const currentWork = [...active.map((item) => ({ type: "Project", title: item.branchName, meta: `${item.locationName}, ${item.city}`, status: item.status, branch: item })), ...openFaults.map((item) => ({ type: "Fault", title: item.title, meta: branches.find((branch) => branch.id === item.branchId)?.branchName || "Unknown Branch", status: item.status }))].slice(0, 8);
  const stats = [
    ["Assigned Projects", filteredProjects.length, "Matched project records"],
    ["Completed Projects", completed.length, "Live or fully handed over"],
    ["Active Projects", active.length, "Planning or in progress"],
    ["Faults Attended", filteredFaults.length, "Recorded fault activity"],
    ["Faults Closed", closedFaults.length, "Resolved records"],
    ["Open Faults", openFaults.length, "Pending follow-up"],
    ["Handovers", filteredHandovers.length, "Completed or involved"],
    ["Pullouts", filteredPullouts.length, "Removed or involved"],
    ["CCTV Reviews", filteredReviews.length, "NVR review records"],
  ];
  return <section className="mt-6 space-y-4">
    <div className="professional-page-card rounded-2xl p-5"><div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between"><div><h3 className="text-lg font-bold">User Portfolio</h3><p className="text-sm text-slate-500">Operational work evidence for projects, faults, handovers, pullouts and CCTV reviews.</p></div><div className="flex flex-col gap-2 sm:flex-row"><SelectInput value={period} onChange={setPeriod}><option value="month">This Month</option><option value="quarter">This Quarter</option><option value="year">This Year</option><option value="all">All Time</option></SelectInput><SelectInput value={selectedContributor} onChange={setSelectedContributor}>{profiles.map((profile) => <option key={profile.name} value={profile.name}>{profile.name}</option>)}</SelectInput></div></div></div>
    <div className="grid gap-4 xl:grid-cols-[360px_1fr]"><div className="professional-page-card rounded-2xl p-5"><div className="flex items-center gap-4"><div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-slate-900 text-xl font-bold text-white">{selectedProfile.name.slice(0,1)}</div><div><h3 className="text-2xl font-bold">{selectedProfile.name}</h3><p className="text-sm text-slate-500">{selectedProfile.roleTitle}</p></div></div><div className="mt-5 grid gap-3 text-sm"><p><b>Company:</b> {selectedProfile.company}</p><p><b>Department:</b> {selectedProfile.department}</p><p><b>Reporting To:</b> {selectedProfile.reportingTo}</p><p><b>Scope:</b> {selectedProfile.locationScope}</p>{selectedProfile.email && <p><b>Email:</b> {selectedProfile.email}</p>}{selectedProfile.mobile && <p><b>Mobile:</b> {selectedProfile.mobile}</p>}</div>{isAdmin && <button onClick={() => openProfile(selectedProfile)} className="mt-5 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Edit Profile</button>}</div><div className="professional-page-card rounded-2xl p-5"><h3 className="text-lg font-bold">Profile Summary</h3><p className="mt-3 text-sm leading-7 text-slate-600">{selectedProfile.profileSummary}</p><div className="mt-5 rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Specialization</p><p className="mt-2 text-sm font-semibold text-slate-900">{selectedProfile.specialization}</p></div></div></div>
    <div className="dashboard-mobile-grid grid gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5">{stats.map(([title, value, note]) => <div key={title} className="professional-page-card rounded-2xl p-4"><p className="text-xs font-semibold text-slate-500">{title}</p><p className="mt-2 text-2xl font-bold text-slate-950">{value}</p><p className="mt-1 text-[11px] text-slate-500">{note}</p></div>)}</div>
    <div className="grid gap-4 xl:grid-cols-2"><div className="professional-page-card rounded-2xl p-5"><h3 className="text-lg font-bold">Current Assigned Work</h3><p className="mt-1 text-sm text-slate-500">Active projects and open faults in the selected period.</p><div className="mt-4 space-y-2">{currentWork.length === 0 ? <p className="rounded-2xl border border-dashed border-slate-300 p-4 text-sm text-slate-500">No active work found.</p> : currentWork.map((item, index) => <button key={`${item.type}-${index}`} onClick={() => item.branch && openBranch(item.branch)} className="flex w-full items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-3 text-left"><div><p className="text-xs font-semibold uppercase text-slate-400">{item.type}</p><p className="font-semibold text-slate-900">{item.title}</p><p className="text-xs text-slate-500">{item.meta}</p></div><Badge type={item.status === "Closed" || item.status === "Live" ? "success" : "warning"}>{item.status}</Badge></button>)}</div></div><div className="professional-page-card rounded-2xl p-5"><h3 className="text-lg font-bold">Recent Project Evidence</h3><p className="mt-1 text-sm text-slate-500">Latest project records linked to {selectedContributor}.</p><div className="mt-4 space-y-2">{filteredProjects.slice(0,8).map((branch) => <button key={branch.id} onClick={() => openBranch(branch)} className="flex w-full items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-3 text-left"><div><p className="font-semibold text-slate-900">{branch.branchName}</p><p className="text-xs text-slate-500">{branch.locationName}, {branch.city}</p></div><Badge type={branch.status === "Live" ? "success" : "warning"}>{branch.status}</Badge></button>)}{filteredProjects.length === 0 && <p className="rounded-2xl border border-dashed border-slate-300 p-4 text-sm text-slate-500">No linked project records.</p>}</div></div></div>
    <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-xs leading-5 text-amber-800">Legacy portfolio totals still rely on names stored in Prepared By, Attended By and similar fields. A later user-ID migration will make these statistics fully reliable.</div>
  </section>;
}
function GroupsSection({ groups, branches, assets, isAdmin, openModal }) {
  return <section className="mt-4 space-y-4">
    <div><h2 className="text-2xl font-bold tracking-tight text-slate-950">Groups & Brands</h2><p className="mt-1 text-sm text-slate-500">Compact overview of business groups, concepts and branch branding.</p></div>
    <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">{groups.map((group) => {
      const groupBranches = branches.filter((branch) => branch.groupName === group.name);
      const branchIds = groupBranches.map((branch) => branch.id);
      const groupAssets = assets.filter((asset) => branchIds.includes(asset.branchId));
      const cameraCount = groupAssets.filter((asset) => asset.category === "CCTV Camera").reduce((sum, asset) => sum + Math.max(1, Number(asset.quantity || 1)), 0);
      const concepts = Array.isArray(group.concepts) ? group.concepts : [];
      return <div key={group.id} className="professional-page-card flex min-h-[250px] flex-col rounded-2xl p-4">
        <div className="flex items-start justify-between gap-2">
          <div className="flex min-w-0 items-center gap-2.5"><div className="flex h-10 w-12 shrink-0 items-center justify-center rounded-lg border border-slate-200 bg-slate-50 p-1">{group.logoDataUrl ? <img src={group.logoDataUrl} alt={`${group.name} logo`} className="max-h-8 max-w-full object-contain" /> : <span className="text-[10px] font-bold text-slate-400">LOGO</span>}</div><div className="min-w-0"><h3 className="truncate text-base font-bold text-slate-950">{group.name}</h3><p className="mt-0.5 text-[11px] text-slate-500">Business group</p></div></div>
          <span className="shrink-0 rounded-full bg-blue-100 px-2 py-1 text-[10px] font-semibold text-blue-700">{groupBranches.length} {groupBranches.length === 1 ? "Branch" : "Branches"}</span>
        </div>
        <div className="mt-3"><p className="text-[10px] font-semibold uppercase tracking-wide text-slate-400">Concepts</p><div className="mt-1.5 flex max-h-[54px] flex-wrap gap-1.5 overflow-hidden">{concepts.length ? concepts.map((concept) => <span key={concept} className="rounded-full bg-purple-100 px-2 py-1 text-[10px] font-semibold text-purple-700">{concept}</span>) : <span className="text-xs text-slate-400">No concepts</span>}</div></div>
        <div className="mt-3 grid grid-cols-2 gap-2"><div className="rounded-xl bg-slate-50 p-2.5"><p className="text-lg font-bold text-slate-950">{groupAssets.reduce((sum, asset) => sum + Math.max(1, Number(asset.quantity || 1)), 0)}</p><p className="text-[10px] text-slate-500">Assets</p></div><div className="rounded-xl bg-slate-50 p-2.5"><p className="text-lg font-bold text-slate-950">{cameraCount}</p><p className="text-[10px] text-slate-500">Cameras</p></div></div>
        <div className="mt-3 space-y-1 border-t border-slate-200 pt-3 text-[10px] text-slate-500"><p className="truncate"><strong className="text-slate-700">Phone:</strong> {group.footerPhone || "Not updated"}</p><p className="truncate"><strong className="text-slate-700">Website:</strong> {group.footerWebsite || "Not updated"}</p><div className="mt-2 h-1.5 rounded-full" style={{ backgroundColor: group.footerColor || "#991b1e" }} /></div>
        {isAdmin && <button onClick={() => openModal("groupBranding", group)} className="mt-auto pt-3 text-left text-xs font-semibold text-blue-600 hover:text-blue-700">Edit branding →</button>}
      </div>;
    })}</div>
  </section>;
}

function LocationsSection({ locations, branches, isAdmin, search, setSearch, openAdd, openEdit, openImport, openTemplate, deleteLocation }) {
  const safeLocations = Array.isArray(locations) ? locations.filter(Boolean) : [];
  const safeBranches = Array.isArray(branches) ? branches.filter(Boolean) : [];
  const query = normalizeExcelText(search || "");
  const filteredLocations = safeLocations.filter((location) => {
    const locName = location.name || location.locationName || "";
    const locCity = location.city || "";
    const locEmirate = location.emirate || "";
    const locType = location.type || "Mall";
    const locNotes = location.notes || "";
    if (!query) return true;
    return normalizeExcelText(`${locName} ${locCity} ${locEmirate} ${locType} ${locNotes}`).includes(query);
  });
  return <section className="mt-6 space-y-4">
    <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div>
          <h3 className="text-lg font-bold">Locations</h3>
          <p className="text-sm text-slate-500">Mall/location master list. Showing {filteredLocations.length} of {safeLocations.length} location(s).</p>
        </div>
        {isAdmin && <div className="flex flex-wrap gap-2"><button onClick={openTemplate} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Download Location Template</button><button onClick={openImport} className="rounded-2xl border border-green-200 bg-green-50 px-4 py-2 text-sm font-semibold text-green-700">Import Locations Excel</button><button onClick={openAdd} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add Location</button></div>}
      </div>
      <div className="mt-4"><TextInput value={search || ""} onChange={setSearch} placeholder="Search by location, city, emirate, type or notes..." /></div>
    </div>
    <Table headers={isAdmin ? ["Location", "City", "Emirate", "Type", "Branches", "Actions"] : ["Location", "City", "Emirate", "Type", "Branches"]} rows={filteredLocations} renderDesktop={(location) => {
      const usedCount = safeBranches.filter((branch) => branch.locationName === (location.name || location.locationName || "") || normalizeLocationKey(branch.locationName, branch.city, branch.emirate) === normalizeLocationKey(location.name || location.locationName || "", location.city, location.emirate)).length;
      return <tr key={location.id || `${location.name}-${location.city}-${location.emirate}`} className="hover:bg-slate-50"><td className="px-4 py-3 font-semibold">{location.name || location.locationName || "Unnamed Location"}<p className="text-xs font-normal text-slate-500">{location.notes || ""}</p></td><td className="px-4 py-3 text-slate-600">{location.city || "-"}</td><td className="px-4 py-3 text-slate-600">{location.emirate || "-"}</td><td className="px-4 py-3"><Badge>{location.type || "Mall"}</Badge></td><td className="px-4 py-3 text-slate-600">{usedCount}</td>{isAdmin && <td className="px-4 py-3"><div className="flex flex-wrap gap-2"><button onClick={() => openEdit(location)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteLocation(location)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div></td>}</tr>;
    }} renderMobile={(location) => {
      const usedCount = safeBranches.filter((branch) => branch.locationName === (location.name || location.locationName || "") || normalizeLocationKey(branch.locationName, branch.city, branch.emirate) === normalizeLocationKey(location.name || location.locationName || "", location.city, location.emirate)).length;
      return <MobileCard key={location.id || `${location.name}-${location.city}-${location.emirate}`} kicker="Location" title={location.name || location.locationName || "Unnamed Location"} subtitle={`${location.city || "-"}, ${location.emirate || "-"}`} status={<Badge>{location.type || "Mall"}</Badge>} details={[["Branches", usedCount], ["Notes", location.notes || "-"]]} actions={isAdmin && <div className="flex gap-2"><button onClick={() => openEdit(location)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteLocation(location)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div>} />;
    }} />
  </section>;
}

function BranchesSection({ branches, assets, search, setSearch, openBranch, isAdmin, openImport, openTemplate }) {
  return <section className="mt-6 space-y-4"><div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between"><div><h3 className="text-lg font-bold">Branches</h3><p className="text-sm text-slate-500">Branch details, IT assets and network setup are inside each branch.</p></div><div className="flex w-full flex-col gap-2 md:w-auto md:flex-row md:items-center">{isAdmin && <><button onClick={openTemplate} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Download Branch Template</button><button onClick={openImport} className="rounded-2xl border border-green-200 bg-green-50 px-4 py-2 text-sm font-semibold text-green-700">Import Branches Excel</button></>}<input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search by branch, location, brand, city..." className="w-full rounded-2xl border border-slate-300 px-4 py-2 text-sm outline-none md:w-96" /></div></div></div><Table headers={["Completed", "Group", "Concept", "Branch", "Location", "Status", "Open"]} rows={branches} renderDesktop={(b) => { const c = assets.filter((a) => a.branchId === b.id).length; return <tr key={b.id} className="hover:bg-slate-50"><td className="px-4 py-3 text-slate-600">{b.completedDate || "Pending"}</td><td className="px-4 py-3 text-slate-600">{b.groupName}</td><td className="px-4 py-3"><Badge type="purple">{b.conceptName}</Badge></td><td className="px-4 py-3 font-semibold">{b.branchName}<p className="text-xs font-normal text-slate-500">{c} assets</p></td><td className="px-4 py-3 text-slate-600">{b.locationName}, {b.city}</td><td className="px-4 py-3"><Badge type={b.status === "Live" ? "success" : "warning"}>{b.status}</Badge></td><td className="px-4 py-3"><button onClick={() => openBranch(b)} className="rounded-xl border border-slate-300 px-3 py-1.5 text-xs font-semibold">Open Branch</button></td></tr>; }} renderMobile={(b) => <MobileCard key={b.id} kicker="Branch" title={b.branchName} subtitle={`${b.locationName}, ${b.city}`} status={<Badge type={b.status === "Live" ? "success" : "warning"}>{b.status}</Badge>} details={[["Group", b.groupName], ["Concept", <Badge type="purple">{b.conceptName}</Badge>], ["Completed", b.completedDate || "Pending"], ["Assets", assets.filter((a) => a.branchId === b.id).length]]} actions={<button onClick={() => openBranch(b)} className="rounded-xl border border-slate-300 px-3 py-1.5 text-xs font-semibold">Open Branch</button>} />} /></section>;
}

function AssetsSection({ assets, branches, search, setSearch, openBranch, isAdmin, assetProducts = [], assetTemplates = [], allAssets = [], assetCategories = [], openModal, deleteProduct, deleteProductsBulk, setProductsActive, deleteAssetTemplate }) {
  const [view, setView] = useState("register");
  const [productSearch, setProductSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("All");
  const [brandFilter, setBrandFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [sortBy, setSortBy] = useState("product");
  const [selectedIds, setSelectedIds] = useState([]);
  const activeCategories = assetCategories.filter((item) => item.isActive !== false).map((item) => item.name);
  const productCategories = Array.from(new Set([...activeCategories, ...assetProducts.map((item) => item.category).filter(Boolean)])).sort();
  const brands = Array.from(new Set(assetProducts.map((item) => item.brand).filter(Boolean))).sort();
  const filteredProducts = useMemo(() => {
    const keyword = productSearch.trim().toLowerCase();
    const result = assetProducts.filter((product) => {
      const matchesSearch = !keyword || `${product.productName} ${product.category} ${product.brand} ${product.model} ${product.description} ${product.defaultLocation} ${product.remarks}`.toLowerCase().includes(keyword);
      const matchesCategory = categoryFilter === "All" || product.category === categoryFilter;
      const matchesBrand = brandFilter === "All" || product.brand === brandFilter;
      const active = product.isActive !== false;
      const matchesStatus = statusFilter === "All" || (statusFilter === "Active" ? active : !active);
      return matchesSearch && matchesCategory && matchesBrand && matchesStatus;
    });
    return [...result].sort((a,b) => sortBy === "category" ? String(a.category).localeCompare(String(b.category)) : sortBy === "brand" ? String(a.brand).localeCompare(String(b.brand)) : sortBy === "updated" ? String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")) : String(a.productName).localeCompare(String(b.productName)));
  }, [assetProducts, productSearch, categoryFilter, brandFilter, statusFilter, sortBy]);
  const visibleIds = filteredProducts.map((item) => item.id);
  const selectedProducts = assetProducts.filter((item) => selectedIds.includes(item.id));
  const allVisibleSelected = visibleIds.length > 0 && visibleIds.every((id) => selectedIds.includes(id));
  useEffect(() => { setSelectedIds((ids) => ids.filter((id) => assetProducts.some((item) => item.id === id))); }, [assetProducts]);
  useEffect(() => { setProductSearch(""); setCategoryFilter("All"); setBrandFilter("All"); setStatusFilter("All"); setSelectedIds([]); }, [view]);
  function toggle(id) { setSelectedIds((ids) => ids.includes(id) ? ids.filter((item) => item !== id) : [...ids, id]); }
  function toggleAll() { setSelectedIds((ids) => allVisibleSelected ? ids.filter((id) => !visibleIds.includes(id)) : Array.from(new Set([...ids, ...visibleIds]))); }
  function exportProducts(list, label) {
    if (!list.length) return alert("No products available to export.");
    const rows = list.map((product, index) => ({ "#": index + 1, "Product Name": product.productName, "Device Category": product.category, "Brand": product.brand || "", "Model": product.model || "", "Description": product.description || "", "Default Quantity": product.defaultQuantity || 1, "Default Installed Location": product.defaultLocation || "", "Default Status": product.defaultStatus || "", "Product Status": product.isActive === false ? "Inactive" : "Active", "Remarks": product.remarks || "" }));
    const ws = XLSX.utils.json_to_sheet(rows); ws["!cols"] = [{wch:5},{wch:30},{wch:20},{wch:18},{wch:22},{wch:38},{wch:12},{wch:24},{wch:18},{wch:14},{wch:34}];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Product List"); XLSX.writeFile(wb, `Asset Product List - ${label}.xlsx`);
  }
  const categoryCounts = productCategories.map((category) => ({ category, count: assetProducts.filter((item) => item.category === category).length })).filter((item) => item.count > 0);
  return <section className="mt-6 space-y-4">
    <div className="professional-page-card rounded-2xl p-5"><div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between"><div><h3 className="text-lg font-bold">IT Assets</h3><p className="text-sm text-slate-500">Global asset register, product master and standard branch templates.</p></div><div className="flex flex-wrap gap-2">{[["register","Asset Register"],["products","Asset Product List"],["templates","Branch Asset Templates"]].map(([key,label]) => <button key={key} onClick={() => setView(key)} className={`rounded-2xl px-4 py-2 text-sm font-semibold ${view === key ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-700"}`}>{label}</button>)}</div></div></div>
    {view === "register" && <><div className="professional-page-card rounded-2xl p-5"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h4 className="font-bold">Global IT Asset Register</h4><p className="text-sm text-slate-500">Search across all branches by serial number, device, IP, branch, location or brand.</p></div><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search serial number, NVR, IP, location..." className="w-full rounded-2xl border border-slate-300 px-4 py-2 text-sm outline-none md:w-[420px]" /></div></div><Table headers={["Category","Device","Description","Qty","Serial No.","IP","Branch","Location","Brand","Status"]} rows={assets} renderDesktop={(asset) => <tr key={asset.id} className="cursor-pointer hover:bg-slate-50" onClick={() => openBranch(branches.find((branch) => branch.id === asset.branchId))}><td className="px-4 py-3"><Badge>{asset.category}</Badge></td><td className="px-4 py-3 font-semibold">{asset.deviceName}</td><td className="px-4 py-3 text-slate-600">{cleanImportedPlaceholder(asset.description) || "-"}</td><td className="px-4 py-3 text-slate-600">{asset.quantity || 1}</td><td className="px-4 py-3 text-slate-600">{cleanImportedPlaceholder(asset.serialNumber) || "-"}</td><td className="px-4 py-3 text-slate-600">{cleanImportedPlaceholder(asset.ipAddress) || "-"}</td><td className="px-4 py-3 font-semibold">{asset.branch.branchName}</td><td className="px-4 py-3 text-slate-600">{asset.branch.locationName}, {asset.branch.city}</td><td className="px-4 py-3 text-slate-600">{cleanImportedPlaceholder(asset.brand) || "-"}</td><td className="px-4 py-3"><Badge type={asset.status === "Faulty" ? "danger" : "success"}>{asset.status}</Badge></td></tr>} renderMobile={(asset) => <MobileCard key={asset.id} kicker={asset.category} title={asset.deviceName} subtitle={`${asset.branch.branchName} • ${asset.branch.locationName}`} status={<Badge type={asset.status === "Faulty" ? "danger" : "success"}>{asset.status}</Badge>} details={[["Description",cleanImportedPlaceholder(asset.description)||"-"],["Qty",asset.quantity||1],["Serial",cleanImportedPlaceholder(asset.serialNumber)||"-"],["IP",cleanImportedPlaceholder(asset.ipAddress)||"-"],["Brand",cleanImportedPlaceholder(asset.brand)||"-"],["Location",asset.installedLocation]]} onClick={() => openBranch(branches.find((branch) => branch.id === asset.branchId))} />}/></>}
    {view === "products" && <div className="professional-page-card rounded-3xl p-5"><div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between"><div><h4 className="text-lg font-bold">Asset Product List</h4><p className="mt-1 text-sm text-slate-500">Reusable products for smart asset entry and branch templates.</p></div>{isAdmin && <div className="flex flex-wrap gap-2"><button onClick={() => openModal("template", {type:"product"})} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Download Product Template</button><button onClick={() => openModal("productImport")} className="rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-2 text-sm font-semibold text-emerald-700">Upload Product Excel</button><button onClick={() => openModal("product")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add Product</button></div>}</div>
      <div className="mt-4 flex flex-wrap gap-2"><button onClick={() => setCategoryFilter("All")} className={`rounded-full border px-3 py-1.5 text-xs font-semibold ${categoryFilter === "All" ? "border-amber-400 bg-amber-100 text-amber-800" : "border-slate-200 bg-slate-50 text-slate-600"}`}>All Products {assetProducts.length}</button>{categoryCounts.map((item) => <button key={item.category} onClick={() => setCategoryFilter(item.category)} className={`rounded-full border px-3 py-1.5 text-xs font-semibold ${categoryFilter === item.category ? "border-amber-400 bg-amber-100 text-amber-800" : "border-slate-200 bg-slate-50 text-slate-600"}`}>{item.category} {item.count}</button>)}</div>
      <div className="mt-4 grid gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-4 md:grid-cols-2 xl:grid-cols-5"><input value={productSearch} onChange={(event) => setProductSearch(event.target.value)} placeholder="Search product, category, brand, model..." className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm outline-none md:col-span-2"/><SelectInput value={categoryFilter} onChange={setCategoryFilter}><option>All</option>{productCategories.map((category) => <option key={category}>{category}</option>)}</SelectInput><SelectInput value={brandFilter} onChange={setBrandFilter}><option>All</option>{brands.map((brand) => <option key={brand}>{brand}</option>)}</SelectInput><SelectInput value={statusFilter} onChange={setStatusFilter}><option>All</option><option>Active</option><option>Inactive</option></SelectInput><SelectInput value={sortBy} onChange={setSortBy}><option value="product">Sort: Product A-Z</option><option value="category">Sort: Category</option><option value="brand">Sort: Brand</option><option value="updated">Sort: Recently Updated</option></SelectInput></div>
      <div className="mt-4 flex flex-col gap-3 rounded-2xl border border-slate-200 bg-white p-4 lg:flex-row lg:items-center lg:justify-between"><div><p className="text-sm font-bold">Product Actions</p><p className="text-xs text-slate-500">Showing {filteredProducts.length} of {assetProducts.length}. Selected {selectedProducts.length}.</p></div><div className="flex flex-wrap gap-2"><button onClick={toggleAll} className="rounded-xl border border-slate-300 px-3 py-2 text-xs font-semibold">{allVisibleSelected ? "Clear Visible" : "Select All Visible"}</button><button onClick={() => exportProducts(assetProducts,"All")} className="rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2 text-xs font-semibold text-indigo-700">Export All</button><button onClick={() => exportProducts(filteredProducts,"Filtered")} className="rounded-xl border border-blue-200 bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-700">Export Filtered</button>{isAdmin && <button onClick={() => exportProducts(selectedProducts,"Selected")} disabled={!selectedProducts.length} className="rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700 disabled:opacity-40">Export Selected</button>}{isAdmin && <button onClick={() => setProductsActive(selectedIds,false)} disabled={!selectedProducts.length} className="rounded-xl border border-amber-200 bg-amber-50 px-3 py-2 text-xs font-semibold text-amber-700 disabled:opacity-40">Deactivate</button>}{isAdmin && <button onClick={() => deleteProductsBulk(selectedIds)} disabled={!selectedProducts.length} className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-xs font-semibold text-red-700 disabled:opacity-40">Delete Selected</button>}{selectedProducts.length > 0 && <button onClick={() => setSelectedIds([])} className="rounded-xl border border-slate-300 px-3 py-2 text-xs font-semibold">Clear</button>}</div></div>
      <div className="mt-4"><Table headers={isAdmin ? ["Select","Product","Category","Brand","Model","Default Location","Status","Actions"] : ["Product","Category","Brand","Model","Default Location","Status"]} rows={filteredProducts} renderDesktop={(product) => <tr key={product.id} className="hover:bg-slate-50">{isAdmin && <td className="px-4 py-3"><input type="checkbox" checked={selectedIds.includes(product.id)} onChange={() => toggle(product.id)} className="h-4 w-4 accent-slate-900"/></td>}<td className="px-4 py-3 font-semibold">{product.productName}<p className="text-xs font-normal text-slate-500">{product.description || "-"}</p></td><td className="px-4 py-3"><Badge>{product.category}</Badge></td><td className="px-4 py-3 text-slate-600">{product.brand || "-"}</td><td className="px-4 py-3 text-slate-600">{product.model || "-"}</td><td className="px-4 py-3 text-slate-600">{product.defaultLocation || "-"}</td><td className="px-4 py-3"><Badge type={product.isActive === false ? "warning" : "success"}>{product.isActive === false ? "Inactive" : "Active"}</Badge></td>{isAdmin && <td className="px-4 py-3"><div className="flex flex-wrap gap-2"><button onClick={() => openModal("product",product)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button>{product.isActive === false && <button onClick={() => setProductsActive([product.id],true)} className="rounded-lg border border-emerald-200 px-2 py-1 text-xs font-semibold text-emerald-700">Activate</button>}<button onClick={() => deleteProduct(product.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">{product.isActive === false ? "Delete" : "Delete / Deactivate"}</button></div></td>}</tr>} renderMobile={(product) => <MobileCard key={product.id} kicker={product.category} title={product.productName} subtitle={product.description || product.model || "Saved product"} status={<Badge type={product.isActive === false ? "warning" : "success"}>{product.isActive === false ? "Inactive" : "Active"}</Badge>} details={[["Brand",product.brand||"-"],["Model",product.model||"-"],["Default Location",product.defaultLocation||"-"],["Qty",product.defaultQuantity||"1"]]} actions={isAdmin && <div className="flex flex-wrap gap-2"><label className="flex items-center gap-2 rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold"><input type="checkbox" checked={selectedIds.includes(product.id)} onChange={() => toggle(product.id)}/> Select</label><button onClick={() => openModal("product",product)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteProduct(product.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div>}/>} /></div>
    </div>}
    {view === "templates" && <div className="professional-page-card rounded-3xl p-5"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h4 className="text-lg font-bold">Standard Branch Asset Templates</h4><p className="mt-1 text-sm text-slate-500">Create setup templates for QSR, DTF, Cafe, Accommodation, Warehouse or Central Kitchen.</p></div>{isAdmin && <div className="flex flex-wrap gap-2"><button onClick={() => openModal("template", {type:"branchAssetTemplate"})} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Download Template Excel</button><button onClick={() => openModal("assetTemplate")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Create Branch Template</button></div>}</div><div className="mt-4"><Table headers={isAdmin ? ["Template","Branch Type","Items","Description","Actions"] : ["Template","Branch Type","Items","Description"]} rows={assetTemplates} renderDesktop={(template) => <tr key={template.id} className="hover:bg-slate-50"><td className="px-4 py-3 font-semibold">{template.name}</td><td className="px-4 py-3"><Badge>{template.branchType || "Branch"}</Badge></td><td className="px-4 py-3">{(template.rows||[]).length}</td><td className="px-4 py-3 text-slate-600">{template.description || "-"}</td>{isAdmin && <td className="px-4 py-3"><div className="flex gap-2"><button onClick={() => openModal("assetTemplate",template)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteAssetTemplate(template.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div></td>}</tr>} renderMobile={(template) => <MobileCard key={template.id} kicker="Branch Template" title={template.name} subtitle={template.description||"Standard setup"} status={<Badge>{template.branchType||"Branch"}</Badge>} details={[["Items",(template.rows||[]).length]]}/>} /></div></div>}
  </section>;
}
function CctvSection({ records, assets, isAdmin, search, setSearch, openBranch, branches, openCctv }) {
  const [quickFilter, setQuickFilter] = useState("all");
  const [scopeFilter, setScopeFilter] = useState("all");
  const isBlank = (v) => !v || String(v).trim() === "" || ["not updated", "not assigned", "to be updated"].includes(String(v).trim().toLowerCase());
  const parseDate = (v) => { const d = new Date(v); return Number.isNaN(d.getTime()) ? null : d; };
  const isReviewedThisWeek = (n) => { const d = parseDate(n.lastRecordingReviewDate); if (!d) return false; const now = new Date(); const diffDays = (now.setHours(0, 0, 0, 0) - d.setHours(0, 0, 0, 0)) / (1000 * 60 * 60 * 24); return diffDays >= 0 && diffDays <= 7; };
  const classifyBranch = (branch) => {
    const type = String(branch?.branchType || "").toLowerCase();
    const text = `${branch?.groupName || ""} ${branch?.conceptName || ""} ${branch?.branchName || ""}`.toLowerCase();
    if (text.includes("din tai fung") || /\bdtf\b/.test(text)) return "dtf";
    if (type === "office" || type.includes("head office") || text.includes("head office") || /\bho\b/.test(text)) return "office";
    if (type === "qsr" || type.includes("quick service")) return "qsr";
    return "other";
  };
  const branchById = (id) => branches.find((b) => b.id === id);
  const cameraAssets = assets.filter((a) => String(a.category || "").toLowerCase().includes("camera"));
  const cameraCount = (scope = "all") => cameraAssets.filter((a) => scope === "all" || classifyBranch(branchById(a.branchId)) === scope).reduce((sum, a) => sum + Math.max(1, Number(a.quantity || 1)), 0);
  const cameraCounts = { all: cameraCount("all"), qsr: cameraCount("qsr"), dtf: cameraCount("dtf"), office: cameraCount("office") };
  const scopeRecords = records.filter((n) => scopeFilter === "all" || classifyBranch(n.branch || branchById(n.branchId)) === scopeFilter);
  const below60 = scopeRecords.filter((n) => Number(n.recordingDays || 0) < 60);
  const reviewedThisWeek = scopeRecords.filter(isReviewedThisWeek);
  const reviewPending = scopeRecords.filter((n) => !isReviewedThisWeek(n));
  const filteredRecords = scopeRecords.filter((n) => {
    if (quickFilter === "below60") return Number(n.recordingDays || 0) < 60;
    if (quickFilter === "pending") return !isReviewedThisWeek(n);
    if (quickFilter === "week") return isReviewedThisWeek(n);
    return true;
  });
  const getStatus = (n) => { const days = Number(n.recordingDays || 0); if (!days) return { label: "Pending", type: "warning" }; if (days < 60) return { label: "Below 60", type: "danger" }; if (n.reviewStatus && n.reviewStatus !== "OK") return { label: n.reviewStatus, type: "warning" }; return { label: "OK", type: "success" }; };
  const exportRows = (rows, label) => { if (!rows.length) return alert("No NVR records found for export."); const data = rows.map((n) => ({ Scope: classifyBranch(n.branch || branchById(n.branchId)).toUpperCase(), Branch: n.branch?.branchName || "", Location: n.branch?.locationName || "", City: n.branch?.city || "", "NVR Name": n.deviceName || "", "Serial Number": n.serialNumber || "", "IP Address": n.ipAddress || "", Username: n.nvrUsername || "admin", Password: isAdmin ? (n.nvrPassword || "") : "Hidden", "Recording Days": n.recordingDays || "", "Last Review Date": n.lastRecordingReviewDate || "", "Reviewed By": n.reviewedBy || "", "Recording Start Date": n.recordingStartDate || "", "Recording End Date": n.recordingEndDate || "", "Review Status": n.reviewStatus || (Number(n.recordingDays || 0) < 60 ? "Below 60 Days" : "OK"), "Issue Found": n.issueFound || "", "Action Required": n.actionRequired || "", "Next Review Date": n.nextReviewDate || "", Remarks: n.remarks || "" })); const wb = XLSX.utils.book_new(); const ws = XLSX.utils.json_to_sheet(data); ws["!cols"] = Object.keys(data[0] || {}).map((key) => ({ wch: Math.max(14, key.length + 2) })); XLSX.utils.book_append_sheet(wb, ws, "NVR Review"); XLSX.writeFile(wb, `CCTV_NVR_${label}_${new Date().toISOString().slice(0, 10)}.xlsx`); };
  const scopeOptions = [
    { id: "all", label: "All", count: cameraCounts.all },
    { id: "qsr", label: "QSR", count: cameraCounts.qsr },
    { id: "dtf", label: "DTF", count: cameraCounts.dtf },
    { id: "office", label: "Office", count: cameraCounts.office },
  ];
  const selectedScopeLabel = scopeOptions.find((item) => item.id === scopeFilter)?.label || "All";
  const activeStatusLabel = quickFilter === "below60" ? "Below 60 Days" : quickFilter === "pending" ? "Review Pending" : quickFilter === "week" ? "Reviewed This Week" : "All NVR";
  return <section className="mt-4 space-y-4">
    <div className="professional-page-card rounded-2xl p-4"><div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between"><div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-xl bg-blue-50 text-blue-600"><AppIcon name="cctv" className="h-4 w-4" /></div><div><h3 className="text-lg font-bold text-slate-950">Weekly Review Register</h3><p className="mt-1 text-sm text-slate-500">Camera deployment counts and weekly NVR recording review by operation type.</p></div></div><div className="flex w-full flex-col gap-2 md:w-auto md:flex-row md:items-center"><button onClick={() => exportRows(filteredRecords, `${scopeFilter}_${quickFilter}`)} className="rounded-xl bg-emerald-50 px-4 py-2.5 text-sm font-semibold text-emerald-700 ring-1 ring-emerald-200">Export Current View</button><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search branch, location, NVR, serial or IP..." className="w-full rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm outline-none md:w-[360px]" /></div></div></div>
    <div className="professional-page-card rounded-2xl p-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div className="flex items-center gap-3"><span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-blue-600"><AppIcon name="cctv" className="h-4 w-4" /></span><div><h3 className="text-sm font-bold text-slate-950">Camera Deployment</h3><p className="mt-0.5 text-xs text-slate-500">Choose an operation to update both camera count and NVR records.</p></div></div><p className="text-xs font-semibold text-slate-500">Selected: <span className="text-slate-900">{selectedScopeLabel}</span></p></div><div className="mt-3 grid grid-cols-2 gap-2 md:grid-cols-4">{scopeOptions.map((item) => <button key={item.id} onClick={() => setScopeFilter(item.id)} className={`flex items-center justify-between rounded-xl border px-3 py-2.5 text-left transition ${scopeFilter === item.id ? "border-blue-500 bg-blue-50 ring-2 ring-blue-100" : "border-slate-200 bg-slate-50 hover:border-blue-300"}`}><span><span className="block text-[10px] font-semibold uppercase tracking-wide text-slate-400">{item.label}</span><span className="mt-0.5 block text-lg font-bold text-slate-950">{item.count}</span></span><span className={`h-2.5 w-2.5 rounded-full ${scopeFilter === item.id ? "bg-blue-600" : "bg-slate-300"}`} /></button>)}</div></div>
    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4"><Card title="Total NVR" value={scopeRecords.length} note={`${selectedScopeLabel} operation`} icon="cctv" accent="blue" onClick={() => setQuickFilter("all")} active={quickFilter === "all"} /><Card title="Below 60 Days" value={below60.length} note="Requires recording follow-up" icon="faults" accent="rose" onClick={() => setQuickFilter("below60")} active={quickFilter === "below60"} /><Card title="Review Pending" value={reviewPending.length} note="Not reviewed in last 7 days" icon="calendar" accent="amber" onClick={() => setQuickFilter("pending")} active={quickFilter === "pending"} /><Card title="Reviewed This Week" value={reviewedThisWeek.length} note="Updated during last 7 days" icon="check" accent="green" onClick={() => setQuickFilter("week")} active={quickFilter === "week"} /></div>
    <div className="flex flex-col gap-1 px-1 sm:flex-row sm:items-center sm:justify-between"><p className="text-xs text-slate-500">Showing <span className="font-semibold text-slate-900">{filteredRecords.length}</span> record(s) for {selectedScopeLabel} • {activeStatusLabel}</p>{quickFilter !== "all" && <button onClick={() => setQuickFilter("all")} className="text-xs font-semibold text-blue-600 hover:text-blue-700">Clear status filter</button>}</div>
    <Table headers={["Branch / Location", "Operation", "NVR / Serial", "IP Address", "Login", "Recording", "Last Review", "Status", "Actions"]} rows={filteredRecords} renderDesktop={(n) => { const status = getStatus(n); const low = Number(n.recordingDays || 0) < 60; const pending = !isReviewedThisWeek(n); const scope = classifyBranch(n.branch); return <tr key={n.id} className={low ? "bg-red-50 hover:bg-red-100" : pending ? "bg-amber-50 hover:bg-amber-100" : "hover:bg-slate-50"}><td className="px-4 py-3 font-semibold">{n.branch.branchName}<p className="text-xs font-normal text-slate-500">{n.branch.locationName}, {n.branch.city}</p>{n.actionRequired && <p className="mt-1 text-xs font-semibold text-red-700">Action: {n.actionRequired}</p>}</td><td className="px-4 py-3"><Badge type={scope === "dtf" ? "purple" : scope === "qsr" ? "success" : scope === "office" ? "warning" : "default"}>{scope.toUpperCase()}</Badge></td><td className="px-4 py-3">{n.deviceName}<p className="text-xs text-slate-500">{n.serialNumber}</p></td><td className="px-4 py-3 text-slate-600">{n.ipAddress || "Not assigned"}</td><td className="px-4 py-3 text-slate-600">{n.nvrUsername || "admin"}<p className="text-xs text-slate-500">{isAdmin ? n.nvrPassword || "To be updated" : "Hidden"}</p></td><td className="px-4 py-3"><Badge type={low ? "danger" : "success"}>{n.recordingDays || 0} days</Badge><p className="mt-1 text-xs text-slate-500">{n.recordingStartDate && n.recordingEndDate ? `${n.recordingStartDate} to ${n.recordingEndDate}` : "Set start/end"}</p></td><td className="px-4 py-3 text-slate-600">{n.lastRecordingReviewDate || "Not updated"}<p className="text-xs text-slate-500">{n.reviewedBy || "Not assigned"}</p></td><td className="px-4 py-3"><Badge type={status.type}>{status.label}</Badge></td><td className="px-4 py-3"><div className="flex flex-wrap gap-2"><button onClick={() => openBranch(branches.find((b) => b.id === n.branchId))} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Open</button>{isAdmin && <button onClick={() => openCctv(n)} className="rounded-lg bg-blue-600 px-2 py-1 text-xs font-semibold text-white">Review / Update</button>}</div></td></tr>; }} renderMobile={(n) => { const status = getStatus(n); const low = Number(n.recordingDays || 0) < 60; const scope = classifyBranch(n.branch); return <MobileCard key={n.id} kicker={`${scope.toUpperCase()} • Branch / Location`} title={n.branch.branchName} subtitle={`${n.branch.locationName}, ${n.branch.city}`} status={<Badge type={status.type}>{status.label}</Badge>} danger={low} details={[["NVR", <>{n.deviceName}<p className="text-xs text-slate-500">{n.serialNumber}</p></>], ["IP", n.ipAddress || "Not assigned"], ["Login", <>{n.nvrUsername || "admin"}<p className="text-xs text-slate-500">{isAdmin ? n.nvrPassword || "To be updated" : "Hidden"}</p></>], ["Recording", <Badge type={low ? "danger" : "success"}>{n.recordingDays || 0} days</Badge>], ["Last Review", n.lastRecordingReviewDate || "Not updated"], ["Reviewed By", n.reviewedBy || "-"]]} actions={<div className="flex gap-2"><button onClick={() => openBranch(branches.find((b) => b.id === n.branchId))} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Open</button>{isAdmin && <button onClick={() => openCctv(n)} className="rounded-lg bg-blue-600 px-2 py-1 text-xs font-semibold text-white">Review</button>}</div>} />; }} />
  </section>;
}

function FaultsSection({ faults, isAdmin, search, setSearch, faultView, setFaultView, range, setRange, openModal, openBranch, deleteFault, downloadCSV }) {
  const viewRows = faults;
  return <section className="mt-6 space-y-4"><div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h3 className="text-lg font-bold">Global Fault Logs</h3><p className="text-sm text-slate-500">Overall, weekly and monthly view for all branch faults.</p></div><div className="flex w-full flex-col gap-2 md:w-auto md:flex-row md:items-center">{isAdmin && <button onClick={() => openModal("fault")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add Fault Log</button>}<input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search branch, fault, attended by, status..." className="w-full rounded-2xl border border-slate-300 px-4 py-2 text-sm outline-none md:w-[420px]" /></div></div></div><div className="dashboard-mobile-grid grid gap-4 md:grid-cols-3 xl:grid-cols-6"><Card title="Overall Tickets" value={faults.length} note="All fault logs" /><Card title="Current View" value={viewRows.length} note={faultView} /><Card title="Open" value={viewRows.filter((f) => f.status === "Open").length} note="Need action" /><Card title="In Progress" value={viewRows.filter((f) => ["In Progress", "Waiting Vendor", "Waiting Parts"].includes(f.status)).length} note="Under follow-up" /><Card title="Closed" value={viewRows.filter((f) => f.status === "Closed").length} note="Resolved" /><Card title="Report Rows" value={viewRows.length} note="Selected range" /></div><div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="grid gap-4 xl:grid-cols-2"><div><p className="text-sm font-semibold text-slate-700">Ticket View</p><div className="mt-2 flex flex-wrap gap-2">{["overall", "weekly", "monthly"].map((k) => <button key={k} onClick={() => setFaultView(k)} className={`rounded-2xl px-4 py-2 text-sm font-semibold ${faultView === k ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-700"}`}>{k[0].toUpperCase() + k.slice(1)}</button>)}</div></div><div><p className="text-sm font-semibold text-slate-700">Download Report by Date</p><div className="mt-2 grid gap-2 md:grid-cols-3"><TextInput type="date" value={range.start} onChange={(v) => setRange((x) => ({ ...x, start: v }))} /><TextInput type="date" value={range.end} onChange={(v) => setRange((x) => ({ ...x, end: v }))} /><button onClick={downloadCSV} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Download CSV</button></div></div></div></div><Table headers={isAdmin ? ["Date/Time", "Branch / Location", "Category", "Fault", "Attended By", "Action Taken", "Status", "Actions"] : ["Date/Time", "Branch / Location", "Category", "Fault", "Attended By", "Action Taken", "Status"]} rows={viewRows} renderDesktop={(f) => <tr key={f.id} className={f.status === "Closed" ? "hover:bg-slate-50" : "bg-yellow-50 hover:bg-yellow-100"}><td className="px-4 py-3 text-slate-600">{f.faultDate}<p className="text-xs text-slate-400">{f.faultTime}</p></td><td className="px-4 py-3 font-semibold">{f.branch.branchName}<p className="text-xs font-normal text-slate-500">{f.branch.locationName}, {f.branch.city}</p></td><td className="px-4 py-3"><Badge type={f.category === "CCTV" ? "purple" : "info"}>{f.category}</Badge></td><td className="px-4 py-3 font-semibold">{f.title}<p className="text-xs font-normal text-slate-500">Reported by: {f.reportedBy}</p></td><td className="px-4 py-3 text-slate-600">{f.attendedBy}</td><td className="px-4 py-3 text-slate-600">{f.actionTaken}</td><td className="px-4 py-3"><Badge type={f.status === "Closed" ? "success" : f.status === "In Progress" ? "warning" : "danger"}>{f.status}</Badge></td>{isAdmin && <td className="px-4 py-3"><div className="flex flex-wrap gap-2"><button onClick={() => openBranch(f.branch)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Open</button><button onClick={() => openModal("fault", f)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteFault(f.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div></td>}</tr>} renderMobile={(f) => <MobileCard key={f.id} kicker="Fault" title={f.title} subtitle={`${f.branch.branchName} • ${f.faultDate} ${f.faultTime}`} status={<Badge type={f.status === "Closed" ? "success" : f.status === "In Progress" ? "warning" : "danger"}>{f.status}</Badge>} details={[["Category", <Badge type={f.category === "CCTV" ? "purple" : "info"}>{f.category}</Badge>], ["Attended", f.attendedBy], ["Action", f.actionTaken], ["Reported", f.reportedBy]]} actions={isAdmin && <div className="flex gap-2"><button onClick={() => openBranch(f.branch)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Open</button><button onClick={() => openModal("fault", f)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteFault(f.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div>} />} /></section>;
}

function BranchDetail({ selectedBranch, assets, allAssets, network, handover, pullout, faults, isAdmin, subTab, setSubTab, openModal, deleteBranch, deleteAsset, deleteAssetsBulk, deleteFault, printReport, updateBranchLogo, removeBranchLogo, groupBranding, branchAssetSearch, setBranchAssetSearch, assetTemplates }) {
  const tabs = [["overview", "Overview"], ["assets", "IT Assets"], ["network", "Network Setup"], ["handover", "Handover"], ["pullout", "Pullout"], ["faults", "Fault Log"]];
  return <section className="mt-6 space-y-4"><div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><h3 className="text-2xl font-bold">{selectedBranch.branchName}</h3><p className="mt-1 text-sm text-slate-500">{selectedBranch.locationName}, {selectedBranch.city} • {selectedBranch.branchType}</p><div className="mt-3 flex flex-wrap gap-2"><Badge type="info">{selectedBranch.groupName}</Badge><Badge type="purple">{selectedBranch.conceptName}</Badge><Badge>{selectedBranch.emirate}</Badge></div></div><div className="flex flex-col items-start gap-2 md:items-end"><Badge type={selectedBranch.handover === 100 ? "success" : "warning"}>{selectedBranch.status}</Badge>{isAdmin && <div className="flex flex-wrap gap-2"><button onClick={() => openModal("editBranch")} className="rounded-xl border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700">Edit Branch Details</button><button onClick={deleteBranch} className="rounded-xl border border-red-200 bg-white px-3 py-1.5 text-xs font-semibold text-red-700">Delete Branch</button></div>}</div></div><div className="mt-6 flex flex-wrap gap-2 border-b border-slate-200 pb-4">{tabs.map(([k, l]) => <button key={k} onClick={() => setSubTab(k)} className={`rounded-2xl px-4 py-2 text-sm font-semibold ${subTab === k ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-700"}`}>{l}</button>)}</div>{subTab === "overview" && <BranchOverview b={selectedBranch} assets={allAssets || assets} />}{subTab === "assets" && <BranchAssets selectedBranch={selectedBranch} assets={assets} allAssets={allAssets || assets} totalAssets={(allAssets || assets).length} search={branchAssetSearch} setSearch={setBranchAssetSearch} isAdmin={isAdmin} openModal={openModal} deleteAsset={deleteAsset} deleteAssetsBulk={deleteAssetsBulk} openTemplate={() => openModal("template", { type: "asset" })} openApplyTemplate={() => openModal("applyAssetTemplate")} assetTemplates={assetTemplates} />}{subTab === "network" && <NetworkTab network={network} isAdmin={isAdmin} openModal={openModal} />}{subTab === "handover" && <HandoverTab selectedBranch={selectedBranch} assets={assets} network={network} handover={handover} isAdmin={isAdmin} openModal={openModal} printReport={printReport} updateBranchLogo={updateBranchLogo} removeBranchLogo={removeBranchLogo} groupBranding={groupBranding} />}{subTab === "pullout" && <PulloutTab selectedBranch={selectedBranch} assets={assets} pullout={pullout} isAdmin={isAdmin} openModal={openModal} printReport={printReport} updateBranchLogo={updateBranchLogo} removeBranchLogo={removeBranchLogo} groupBranding={groupBranding} />}{subTab === "faults" && <BranchFaultTab faults={faults} isAdmin={isAdmin} openModal={openModal} deleteFault={deleteFault} />}</div></section>;
}
function BranchOverview({ b, assets }) { return <div className="mt-6 space-y-6"><div className="dashboard-mobile-grid grid gap-4 md:grid-cols-2 xl:grid-cols-5">{[["Opening Date", b.openingDate], ["Completed Date", b.completedDate || "Pending"], ["Branch Manager", b.branchManager], ["Contact Number", b.contactNumber], ["Prepared By", b.preparedBy]].map(([l, v]) => <div key={l} className="rounded-2xl bg-slate-50 p-4"><p className="text-sm font-medium text-slate-500">{l}</p><p className="mt-2 break-words text-base font-bold leading-snug text-slate-900">{v}</p></div>)}</div><div className="dashboard-mobile-grid grid gap-4 md:grid-cols-3"><div className="rounded-2xl bg-slate-900 p-5 text-white"><p className="text-sm text-slate-300">Branch Assets</p><p className="mt-2 text-3xl font-bold">{assets.length}</p></div><div className="rounded-2xl bg-slate-900 p-5 text-white"><p className="text-sm text-slate-300">Cameras</p><p className="mt-2 text-3xl font-bold">{assets.filter((a) => a.category === "CCTV Camera").length}</p></div><div className="rounded-2xl bg-slate-900 p-5 text-white"><p className="text-sm text-slate-300">Handover</p><p className="mt-2 text-3xl font-bold">{b.handover}%</p></div></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-sm font-semibold text-slate-700">Notes</p><p className="mt-1 text-sm text-slate-600">{b.notes}</p></div></div>; }
function BranchAssets({ selectedBranch, assets, allAssets, totalAssets, search, setSearch, isAdmin, openModal, deleteAsset, deleteAssetsBulk, openTemplate, openApplyTemplate, assetTemplates = [] }) {
  const [selectedIds, setSelectedIds] = useState([]);
  const visibleIds = useMemo(() => assets.map((a) => a.id), [assets]);
  const selectedAssets = useMemo(() => assets.filter((a) => selectedIds.includes(a.id)), [assets, selectedIds]);
  const allVisibleSelected = visibleIds.length > 0 && visibleIds.every((id) => selectedIds.includes(id));

  useEffect(() => {
    setSelectedIds((ids) => ids.filter((id) => visibleIds.includes(id)));
  }, [visibleIds]);

  function toggleAsset(id) {
    setSelectedIds((ids) => ids.includes(id) ? ids.filter((item) => item !== id) : [...ids, id]);
  }

  function toggleSelectAll() {
    setSelectedIds((ids) => allVisibleSelected ? ids.filter((id) => !visibleIds.includes(id)) : Array.from(new Set([...ids, ...visibleIds])));
  }

  function exportAssetsExcel(list, suffix = "IT Assets") {
    const rows = list.map((a, index) => ({
      "#": index + 1,
      "Category": a.category || "",
      "Device Name / Asset Name": cleanImportedPlaceholder(a.deviceName),
      "Description": cleanImportedPlaceholder(a.description),
      "Quantity": a.quantity || 1,
      "Serial Number": cleanImportedPlaceholder(a.serialNumber),
      "Brand": cleanImportedPlaceholder(a.brand),
      "Model": cleanImportedPlaceholder(a.model),
      "IP Address": cleanImportedPlaceholder(a.ipAddress),
      "MAC Address": cleanImportedPlaceholder(a.macAddress),
      "Installed Location": cleanImportedPlaceholder(a.installedLocation),
      "Switch Port": cleanImportedPlaceholder(a.switchPort),
      "Camera Channel": cleanImportedPlaceholder(a.channelNumber),
      "Camera View / Area": cleanImportedPlaceholder(a.cameraView),
      "Warranty Start": cleanImportedPlaceholder(a.warrantyStart),
      "Warranty End": cleanImportedPlaceholder(a.warrantyEnd),
      "Status": a.status || "",
      "Remarks": cleanImportedPlaceholder(a.remarks),
    }));
    if (rows.length === 0) return alert("No assets available to export.");
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = [
      { wch: 5 }, { wch: 18 }, { wch: 28 }, { wch: 42 }, { wch: 8 }, { wch: 22 },
      { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 20 }, { wch: 22 }, { wch: 16 },
      { wch: 16 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 36 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "IT Assets");
    const branchName = `${selectedBranch?.branchName || "Branch"} ${selectedBranch?.locationName || ""}`.trim();
    const safeName = `${branchName} - ${suffix}`.replace(/[\\/:*?"<>|]/g, "-").slice(0, 120);
    XLSX.writeFile(workbook, `${safeName}.xlsx`);
  }

  function deleteSelected() {
    if (selectedAssets.length === 0) return alert("Select at least one asset to delete.");
    deleteAssetsBulk(selectedAssets.map((asset) => asset.id));
    setSelectedIds([]);
  }

  return <div className="mt-6 space-y-4">
    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
      <div>
        <h3 className="text-lg font-bold">Branch IT Assets</h3>
        <p className="text-sm text-slate-500">Search assets inside this branch by serial, name, model, brand, IP, location or category. Showing {assets.length} of {totalAssets} asset(s).</p>
      </div>
      {isAdmin && <div className="flex flex-wrap gap-2">
        <button onClick={openApplyTemplate} disabled={!assetTemplates.length} className="rounded-2xl border border-purple-200 bg-purple-50 px-4 py-2 text-sm font-semibold text-purple-700 disabled:cursor-not-allowed disabled:opacity-50">Apply Asset Template</button>
        <button onClick={() => openModal("assetImport")} className="rounded-2xl border border-green-200 bg-green-50 px-4 py-2 text-sm font-semibold text-green-700">Import from Excel</button>
        <button onClick={() => openModal("asset")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add IT Asset</button>
      </div>}
    </div>

    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm"><TextInput value={search || ""} onChange={setSearch} placeholder="Search by asset name, serial number, model, brand, IP, location, category..." /></div>

    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div>
          <p className="text-sm font-bold text-slate-800">Asset Actions</p>
          <p className="text-xs text-slate-500">Selected {selectedAssets.length} of {assets.length} visible asset(s).</p>
        </div>
        <div className="flex flex-wrap gap-2">
          {isAdmin && <button onClick={toggleSelectAll} className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700">{allVisibleSelected ? "Clear Selection" : "Select All Visible"}</button>}
          {isAdmin && <button onClick={openTemplate} className="rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2 text-xs font-semibold text-indigo-700">Download Asset Template</button>}
          <button onClick={() => exportAssetsExcel(allAssets || assets, "All Branch IT Assets")} className="rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2 text-xs font-semibold text-indigo-700">Export All Branch Assets</button>
          <button onClick={() => exportAssetsExcel(assets, "Visible IT Assets")} className="rounded-xl border border-blue-200 bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-700">Export Visible Assets</button>
          {isAdmin && <button onClick={() => exportAssetsExcel(selectedAssets, "Selected IT Assets")} disabled={selectedAssets.length === 0} className="rounded-xl border border-green-200 bg-green-50 px-3 py-2 text-xs font-semibold text-green-700 disabled:cursor-not-allowed disabled:opacity-50">Export Selected</button>}
          {isAdmin && <button onClick={deleteSelected} disabled={selectedAssets.length === 0} className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-xs font-semibold text-red-700 disabled:cursor-not-allowed disabled:opacity-50">Delete Selected</button>}
        </div>
      </div>
    </div>

    <Table headers={isAdmin ? ["Select", "Category", "Device", "Description", "Qty", "Serial No.", "IP Address", "Installed Location", "Status", "Actions"] : ["Category", "Device", "Description", "Qty", "Serial No.", "IP Address", "Installed Location", "Status"]} rows={assets} renderDesktop={(a) => <tr key={a.id} className="hover:bg-slate-50">
      {isAdmin && <td className="px-4 py-3"><input type="checkbox" checked={selectedIds.includes(a.id)} onChange={() => toggleAsset(a.id)} className="h-4 w-4 accent-slate-900" /></td>}
      <td className="px-4 py-3"><Badge>{a.category}</Badge></td><td className="px-4 py-3 font-semibold">{a.deviceName}</td><td className="px-4 py-3 text-slate-600">{a.description || "-"}</td><td className="px-4 py-3 text-slate-600">{a.quantity || 1}</td><td className="px-4 py-3 text-slate-600">{a.serialNumber}</td><td className="px-4 py-3 text-slate-600">{a.ipAddress}</td><td className="px-4 py-3 text-slate-600">{a.installedLocation}</td><td className="px-4 py-3"><Badge type={a.status === "Faulty" ? "danger" : "success"}>{a.status}</Badge></td>{isAdmin && <td className="px-4 py-3"><div className="flex gap-2"><button onClick={() => openModal("asset", a)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteAsset(a.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div></td>}</tr>} renderMobile={(a) => <MobileCard key={a.id} kicker={a.category} title={a.deviceName} subtitle={a.serialNumber} status={<Badge type={a.status === "Faulty" ? "danger" : "success"}>{a.status}</Badge>} details={[["Description", a.description || "-"], ["Model", a.model || "-"], ["Qty", a.quantity || 1], ["IP", a.ipAddress], ["Location", a.installedLocation], ["Brand", a.brand]]} actions={isAdmin && <div className="flex flex-wrap gap-2"><label className="flex items-center gap-2 rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold"><input type="checkbox" checked={selectedIds.includes(a.id)} onChange={() => toggleAsset(a.id)} className="accent-slate-900" /> Select</label><button onClick={() => openModal("asset", a)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteAsset(a.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div>} />} />
  </div>;
}
function NetworkTab({ network, isAdmin, openModal }) { const rows = [["ISP", network?.isp], ["Internet Account", network?.internetAccount], ["Bandwidth", network?.bandwidth], ["Router Model", network?.routerModel], ["Router / Gateway IP", network?.routerIp], ["Switch Model", network?.switchModel], ["NVR IP", network?.nvrIp], ["CCTV IP Range", network?.cctvRange], ["POS IP Range", network?.posRange], ["Biometric IP", network?.biometricIp], ["Wi-Fi Name", network?.wifiName], ["Wi-Fi Password", network?.wifiPassword], ["VLAN Details", network?.vlanDetails]]; return <div className="mt-6 space-y-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h3 className="text-lg font-bold">Branch Network Setup</h3><p className="text-sm text-slate-500">Each branch has its own ISP, router, IP ranges, NVR IP and POS network details.</p></div>{isAdmin && <button onClick={() => openModal("network")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Add / Edit Network Setup</button>}</div>{!network ? <div className="rounded-2xl bg-slate-50 p-5 text-sm text-slate-600">No network setup added yet.</div> : <div className="dashboard-mobile-grid grid gap-4 md:grid-cols-2 xl:grid-cols-3">{rows.map(([l, v]) => <div key={l} className="rounded-2xl bg-slate-50 p-4"><p className="text-sm font-semibold text-slate-500">{l}</p><p className="mt-2 break-words text-base font-bold leading-snug text-slate-900">{v || "To be updated"}</p></div>)}<div className="rounded-2xl bg-slate-50 p-4 md:col-span-2 xl:col-span-3"><p className="text-sm text-slate-500">Remarks</p><p className="mt-1 text-sm font-semibold">{network.remarks}</p></div></div>}</div>; }
function HandoverTab({ selectedBranch, assets, network, handover, isAdmin, openModal, printReport, updateBranchLogo, removeBranchLogo, groupBranding }) {
  const [template, setTemplate] = useState("standard");
  const templateLabel = template === "store" ? "Store Handover" : "Department Handover";
  return <div className="mt-6 space-y-4">
    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
      <div>
        <h3 className="text-lg font-bold">Branch Handover Forms</h3>
        <p className="text-sm text-slate-500">Department Handover is the standard internal handover report. Store Handover is a customizable template with branch logo and legal declaration.</p>
      </div>
      <div className="flex flex-wrap gap-2">
        {isAdmin && <button onClick={() => openModal("handover")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Fill / Edit Handover Form</button>}
        <button onClick={() => printReport(templateLabel)} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Download PDF</button>
      </div>
    </div>

    <div className="flex flex-wrap gap-2 rounded-2xl bg-slate-50 p-3">
      <button onClick={() => setTemplate("standard")} className={`rounded-xl px-3 py-2 text-xs font-bold ${template === "standard" ? "bg-slate-900 text-white" : "bg-white text-slate-700"}`}>Department Handover</button>
      <button onClick={() => setTemplate("store")} className={`rounded-xl px-3 py-2 text-xs font-bold ${template === "store" ? "bg-slate-900 text-white" : "bg-white text-slate-700"}`}>Store Handover</button>
    </div>

    {template === "store" && <div className="grid gap-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm md:grid-cols-2">
      <div>
        <h4 className="font-bold">Store Handover Logo</h4>
        <p className="mt-1 text-sm text-slate-500">Upload branch/client logo. It overrides group logo only for this branch.</p>
        {isAdmin && <div className="mt-3 flex flex-wrap gap-2">
          <label className="cursor-pointer rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">
            Upload / Replace Logo
            <input type="file" accept="image/png,image/jpeg,image/jpg,image/webp" className="hidden" onChange={(event) => { updateBranchLogo(event.target.files?.[0]); event.target.value = ""; }} />
          </label>
          {selectedBranch.logoDataUrl && <button onClick={removeBranchLogo} className="rounded-2xl border border-red-200 bg-white px-4 py-2 text-sm font-semibold text-red-700">Remove Logo</button>}
        </div>}
      </div>
      <div className="flex min-h-24 items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-4">
        {selectedBranch.logoDataUrl ? <img src={selectedBranch.logoDataUrl} alt="Store logo" className="max-h-24 max-w-full object-contain" /> : <p className="text-sm font-semibold text-slate-500">No logo uploaded</p>}
      </div>
    </div>}

    {!handover ? <div className="rounded-2xl bg-slate-50 p-5 text-sm text-slate-600">No handover form completed yet.</div> : template === "store" ? <StoreHandoverReport selectedBranch={selectedBranch} assets={assets} data={handover} groupBranding={groupBranding} /> : <Report type="handover" selectedBranch={selectedBranch} assets={assets} network={network} data={handover} />}
  </div>;
}
function PulloutTab({ selectedBranch, assets, pullout, isAdmin, openModal, printReport, updateBranchLogo, removeBranchLogo, groupBranding }) {
  const [template, setTemplate] = useState("standard");
  const templateLabel = template === "store" ? "Store Pullout" : "Standard Pullout";

  return <div className="mt-6 space-y-4">
    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
      <div>
        <h3 className="text-lg font-bold">Branch IT Pullout Forms</h3>
        <p className="text-sm text-slate-500">Use when a branch is closed, relocated, or devices are removed from site. Store Pullout uses the same branded footer style as Store Handover.</p>
      </div>
      <div className="flex flex-wrap gap-2">
        {isAdmin && <button onClick={() => openModal("pullout")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Fill / Edit Pullout Form</button>}
        <button onClick={() => printReport(templateLabel)} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Download PDF</button>
      </div>
    </div>

    <div className="flex flex-wrap gap-2 rounded-2xl bg-slate-50 p-3">
      <button onClick={() => setTemplate("standard")} className={`rounded-xl px-3 py-2 text-xs font-bold ${template === "standard" ? "bg-slate-900 text-white" : "bg-white text-slate-700"}`}>Standard Pullout</button>
      <button onClick={() => setTemplate("store")} className={`rounded-xl px-3 py-2 text-xs font-bold ${template === "store" ? "bg-slate-900 text-white" : "bg-white text-slate-700"}`}>Store Pullout</button>
    </div>

    {template === "store" && <div className="grid gap-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm md:grid-cols-2">
      <div>
        <h4 className="font-bold">Store Pullout Logo</h4>
        <p className="mt-1 text-sm text-slate-500">Uses branch logo first, then group logo/footer branding.</p>
        {isAdmin && <div className="mt-3 flex flex-wrap gap-2">
          <label className="cursor-pointer rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">
            Upload / Replace Logo
            <input type="file" accept="image/png,image/jpeg,image/jpg,image/webp" className="hidden" onChange={(event) => { updateBranchLogo(event.target.files?.[0]); event.target.value = ""; }} />
          </label>
          {selectedBranch.logoDataUrl && <button onClick={removeBranchLogo} className="rounded-2xl border border-red-200 bg-white px-4 py-2 text-sm font-semibold text-red-700">Remove Logo</button>}
        </div>}
      </div>
      <div className="flex min-h-24 items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-4">
        {selectedBranch.logoDataUrl ? <img src={selectedBranch.logoDataUrl} alt="Store logo" className="max-h-24 max-w-full object-contain" /> : groupBranding?.logoDataUrl ? <img src={groupBranding.logoDataUrl} alt="Group logo" className="max-h-24 max-w-full object-contain" /> : <p className="text-sm font-semibold text-slate-500">No branch/group logo uploaded</p>}
      </div>
    </div>}

    {!pullout ? <div className="rounded-2xl bg-slate-50 p-5 text-sm text-slate-600">No pullout form completed yet.</div> : template === "store" ? <StorePulloutReport selectedBranch={selectedBranch} assets={assets} data={pullout} groupBranding={groupBranding} /> : <Report type="pullout" selectedBranch={selectedBranch} assets={assets} data={pullout} />}
  </div>;
}
function BranchFaultTab({ faults, isAdmin, openModal, deleteFault }) { return <div className="mt-6 space-y-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h3 className="text-lg font-bold">Branch Fault Log</h3><p className="text-sm text-slate-500">Record branch faults, attended person, action taken and status.</p></div>{isAdmin && <button onClick={() => openModal("fault")} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add Fault Log</button>}</div><Table headers={isAdmin ? ["Date/Time", "Category", "Fault", "Attended By", "Action Taken", "Status", "Actions"] : ["Date/Time", "Category", "Fault", "Attended By", "Action Taken", "Status"]} rows={faults} renderDesktop={(f) => <tr key={f.id} className="hover:bg-slate-50"><td className="px-4 py-3 text-slate-600">{f.faultDate}<p className="text-xs text-slate-400">{f.faultTime}</p></td><td className="px-4 py-3"><Badge type={f.category === "CCTV" ? "purple" : "info"}>{f.category}</Badge></td><td className="px-4 py-3 font-semibold">{f.title}<p className="text-xs font-normal text-slate-500">Reported by: {f.reportedBy}</p></td><td className="px-4 py-3 text-slate-600">{f.attendedBy}</td><td className="px-4 py-3 text-slate-600">{f.actionTaken}</td><td className="px-4 py-3"><Badge type={f.status === "Closed" ? "success" : "danger"}>{f.status}</Badge></td>{isAdmin && <td className="px-4 py-3"><div className="flex gap-2"><button onClick={() => openModal("fault", f)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteFault(f.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div></td>}</tr>} renderMobile={(f) => <MobileCard key={f.id} kicker="Fault" title={f.title} subtitle={`${f.faultDate} ${f.faultTime}`} status={<Badge type={f.status === "Closed" ? "success" : "danger"}>{f.status}</Badge>} details={[["Category", <Badge type={f.category === "CCTV" ? "purple" : "info"}>{f.category}</Badge>], ["Attended", f.attendedBy], ["Action", f.actionTaken], ["Reported", f.reportedBy]]} actions={isAdmin && <div className="flex gap-2"><button onClick={() => openModal("fault", f)} className="rounded-lg border border-slate-300 px-2 py-1 text-xs font-semibold">Edit</button><button onClick={() => deleteFault(f.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></div>} />} /></div>; }



function StorePrintFooter({ footerPhone, footerAddress, footerWebsite, footerColor }) {
  return <div className="store-print-footer" style={{ backgroundColor: footerColor }}>
    <div>☎ &nbsp; {footerPhone}</div>
    <div style={{ textAlign: "center" }}>📍 &nbsp; {footerAddress}</div>
    <div style={{ textAlign: "right" }}>◎ &nbsp; {footerWebsite}</div>
  </div>;
}

function StorePrintHeader({ logo, groupName, footerAddress, footerPhone, footerWebsite }) {
  return <div className="store-print-header">
    <div className="store-print-logo">
      {logo ? <img src={logo} alt="Store logo" /> : <div style={{ fontSize: 8, fontWeight: 800, color: "#64748b", textAlign: "center" }}>STORE<br />LOGO</div>}
    </div>
    <div className="store-print-company">
      <p style={{ margin: 0, fontWeight: 800, color: "#0f172a" }}>{groupName || "Company / Client Name"}</p>
      <p style={{ margin: 0 }}>{footerAddress}</p>
      <p style={{ margin: 0 }}>{footerPhone} • {footerWebsite}</p>
    </div>
  </div>;
}

function StorePrintAssetTable({ rows, title = "Asset Details" }) {
  return <div className="store-print-section">
    <h3>{title}</h3>
    <table className="store-print-table">
      <thead><tr><th style={{ width: "18%" }}>Category</th><th style={{ width: "15%" }}>Model</th><th>Item Description</th><th style={{ width: "16%" }}>Brand</th><th style={{ width: "7%" }}>Qty</th></tr></thead>
      <tbody>{rows.map((a, index) => <tr key={`${a.id || a.serialNumber || index}-${index}`}>
        <td><strong>{a.category || "-"}</strong></td>
        <td>{a.model || "To be updated"}</td>
        <td>{a.description || a.deviceName || "-"}{a.serialNumber && <p style={{ margin: "1mm 0 0", fontSize: 7.6, color: "#64748b" }}>Serial: {a.serialNumber}</p>}</td>
        <td>{a.brand || "To be updated"}</td>
        <td>{a.quantity || 1}</td>
      </tr>)}</tbody>
    </table>
  </div>;
}

function splitStoreAssetRows(rows, firstCount = 9, nextCount = 10) {
  const source = Array.isArray(rows) ? rows : [];
  const chunks = [];
  chunks.push(source.slice(0, firstCount));
  let index = firstCount;
  while (index < source.length) {
    chunks.push(source.slice(index, index + nextCount));
    index += nextCount;
  }
  return chunks.length ? chunks : [[]];
}

function StoreHandoverPrintPages({ selectedBranch, assets, data, logo, footerPhone, footerAddress, footerWebsite, footerColor, declarationItems }) {
  const chunks = splitStoreAssetRows(assets, 9, 10);
  return <div className="store-print-pages hidden">
    {chunks.map((chunk, index) => <div className="store-print-page" key={`handover-assets-${index}`}>
      <StorePrintHeader logo={logo} groupName={selectedBranch.groupName} footerAddress={footerAddress} footerPhone={footerPhone} footerWebsite={footerWebsite} />
      {index === 0 && <>
        <div className="store-print-title"><h2>Store Asset Handover Form</h2><p>New Store Opening – IT / Equipment Assets</p></div>
        <div className="store-print-section">
          <h3>1. Store Information</h3>
          <table className="store-print-table store-print-info"><tbody>
            <tr><td><strong>Store Name</strong></td><td>{selectedBranch.branchName}</td><td><strong>Store Code</strong></td><td></td></tr>
            <tr><td><strong>Location</strong></td><td>{selectedBranch.locationName}</td><td><strong>City</strong></td><td>{selectedBranch.city}</td></tr>
            <tr><td><strong>Date of Handover</strong></td><td>{data.handoverDate}</td><td><strong>Status</strong></td><td>{data.handoverStatus}</td></tr>
          </tbody></table>
        </div>
      </>}
      <StorePrintAssetTable rows={chunk} title={index === 0 ? "2. Asset Details" : "2. Asset Details - Continued"} />
      <StorePrintFooter footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} />
    </div>)}

    <div className="store-print-page">
      <StorePrintHeader logo={logo} groupName={selectedBranch.groupName} footerAddress={footerAddress} footerPhone={footerPhone} footerWebsite={footerWebsite} />
      <div className="store-print-title"><h2>Store Asset Handover Form</h2><p>Confirmation, declaration and signatures</p></div>
      <div className="store-print-section store-print-text">
        <h3>3. Accessories / Additional Items (If Applicable)</h3>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "2mm" }}>
          <p>☐ Power Adapters</p><p>☐ Network Cables</p><p>☐ Tablets</p><p>☐ Mouse / Keyboard</p><p>☐ SIM Card</p><p>☐ Printer Cartridges</p>
        </div>
        <p>☐ Other: ________________________________________________</p>
      </div>
      <div className="store-print-section store-print-text">
        <h3>4. Installation & Testing Confirmation</h3>
        <ul><li>All listed assets have been installed, configured, and tested by IT.</li><li>Devices are fully operational at the time of handover.</li><li>Store representatives have verified physical count, serial numbers, and working condition.</li></ul>
      </div>
      <div className="store-print-section store-print-text">
        <h3>5. Acknowledgement & Legal Responsibility Declaration</h3>
        <p>I, the undersigned Store Manager / Supervisor, hereby acknowledge receipt of the above-mentioned company-owned assets in good working condition. I expressly agree and legally undertake the following obligations:</p>
        <ol>{declarationItems.map((item) => <li key={item}>{item}</li>)}</ol>
      </div>
      <div className="store-print-section">
        <h3>6. Signatures</h3>
        <div className="store-print-signatures">
          <div className="store-print-signbox"><strong>Handed Over by IT Team</strong><p>Name: ______________________________</p><p>Designation: _________________________</p><p>Signature: ___________________________</p><p>Date: ____ / ____ / ______</p></div>
          <div className="store-print-signbox"><strong>Received By (Store Manager / Supervisor)</strong><p>Name: ______________________________</p><p>Employee ID: _________________________</p><p>Designation: _________________________</p><p>Signature: ___________________________</p><p>Date: ____ / ____ / ______</p></div>
        </div>
      </div>
      <StorePrintFooter footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} />
    </div>
  </div>;
}

function StorePulloutPrintPages({ selectedBranch, assets, data, logo, footerPhone, footerAddress, footerWebsite, footerColor }) {
  const chunks = splitStoreAssetRows(assets, 9, 10);
  return <div className="store-print-pages hidden">
    {chunks.map((chunk, index) => <div className="store-print-page" key={`pullout-assets-${index}`}>
      <StorePrintHeader logo={logo} groupName={selectedBranch.groupName} footerAddress={footerAddress} footerPhone={footerPhone} footerWebsite={footerWebsite} />
      {index === 0 && <>
        <div className="store-print-title"><h2>Store Asset Pullout Form</h2><p>Branch Closure / Relocation – IT / Equipment Assets</p></div>
        <div className="store-print-section">
          <h3>1. Store / Pullout Information</h3>
          <table className="store-print-table store-print-info"><tbody>
            <tr><td><strong>Store Name</strong></td><td>{selectedBranch.branchName}</td><td><strong>Store Code</strong></td><td></td></tr>
            <tr><td><strong>Location</strong></td><td>{selectedBranch.locationName}</td><td><strong>City</strong></td><td>{selectedBranch.city}</td></tr>
            <tr><td><strong>Pullout Date</strong></td><td>{data.pulloutDate}</td><td><strong>Status</strong></td><td>{data.pulloutStatus}</td></tr>
            <tr><td><strong>Closure Date</strong></td><td>{data.closureDate || "Not updated"}</td><td><strong>Destination</strong></td><td>{data.destination}</td></tr>
          </tbody></table>
        </div>
      </>}
      <StorePrintAssetTable rows={chunk} title={index === 0 ? "2. Pulled Out Asset Details" : "2. Pulled Out Asset Details - Continued"} />
      <StorePrintFooter footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} />
    </div>)}

    <div className="store-print-page">
      <StorePrintHeader logo={logo} groupName={selectedBranch.groupName} footerAddress={footerAddress} footerPhone={footerPhone} footerWebsite={footerWebsite} />
      <div className="store-print-title"><h2>Store Asset Pullout Form</h2><p>Verification, remarks and signatures</p></div>
      <div className="store-print-section store-print-text">
        <h3>3. Pullout Summary</h3>
        <p><strong>Store Manager:</strong> {data.storeManagerName}</p>
        <p><strong>Removed By:</strong> {data.removedBy}</p>
        <p><strong>In-charge:</strong> {data.inchargeName} - {data.inchargeRole}</p>
        <p><strong>Received By:</strong> {data.receivedBy}</p>
        <p><strong>Pullout Reason:</strong> {data.pulloutReason}</p>
        <p><strong>Asset Condition:</strong> {data.assetCondition}</p>
        <p><strong>Destination:</strong> {data.destination}</p>
      </div>
      <div className="store-print-section store-print-text">
        <h3>4. Pullout & Verification Confirmation</h3>
        <ul><li>All listed assets have been removed from the store/site by IT or authorized personnel.</li><li>Physical count and visible asset condition were checked during pullout.</li><li>Assets are handed over to the destination/receiver mentioned in this document.</li><li>Any missing, damaged, or pending item should be recorded in remarks before signing.</li></ul>
      </div>
      <div className="store-print-section store-print-text"><h3>5. Remarks</h3><p><strong>IT Remarks:</strong> {data.itRemarks || "No remarks updated."}</p><p><strong>Operations / Store Remarks:</strong> {data.operationsRemarks || "No remarks updated."}</p></div>
      <div className="store-print-section"><h3>6. Signatures</h3><div className="store-print-signatures"><div className="store-print-signbox"><strong>Removed / Handed Over by IT Team</strong><p>Name: ______________________________</p><p>Designation: _________________________</p><p>Signature: ___________________________</p><p>Date: ____ / ____ / ______</p></div><div className="store-print-signbox"><strong>Received By / Store Manager / Supervisor</strong><p>Name: ______________________________</p><p>Employee ID: _________________________</p><p>Designation: _________________________</p><p>Signature: ___________________________</p><p>Date: ____ / ____ / ______</p></div></div></div>
      <StorePrintFooter footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} />
    </div>
  </div>;
}

function StorePulloutReport({ selectedBranch, assets, data, groupBranding }) {
  const assetRows = assets.length ? assets : [];
  const branding = groupBranding || {};
  const logo = selectedBranch.logoDataUrl || branding.logoDataUrl || "";
  const footerPhone = branding.footerPhone || "Not configured";
  const footerAddress = branding.footerAddress || "Demo Head Office";
  const footerWebsite = branding.footerWebsite || "example.com";
  const footerColor = branding.footerColor || "#991b1e";

  return <div className="rounded-2xl border border-slate-200 bg-slate-100 p-4 shadow-sm">
    <div className="store-handover-screen mx-auto w-full max-w-[794px] bg-white p-8 text-slate-900 shadow-lg">
      <div className="flex items-center justify-between gap-4 border-b-4 border-slate-900 pb-5">
        <div className="flex h-20 w-52 items-center justify-center rounded-xl border border-slate-200 bg-white p-3">
          {logo ? <img src={logo} alt="Store logo" className="max-h-16 max-w-full object-contain" /> : <div className="text-center text-xs font-bold text-slate-500">STORE<br />LOGO</div>}
        </div>
        <div className="text-right text-xs leading-relaxed text-slate-600">
          <p className="font-bold text-slate-900">{selectedBranch.groupName || "Company / Client Name"}</p>
          <p>{footerAddress}</p>
          <p>{footerPhone} • {footerWebsite}</p>
        </div>
      </div>

      <div className="mt-8 text-center">
        <h2 className="text-xl font-bold uppercase tracking-wide">Store Asset Pullout Form</h2>
        <p className="mt-1 text-xs text-slate-500">Branch Closure / Relocation – IT / Equipment Assets</p>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">1. Store / Pullout Information</h3>
        <table className="w-full border-collapse text-sm">
          <tbody>
            <tr><td className="border px-2 py-2 font-semibold">Store Name</td><td className="border px-2 py-2">{selectedBranch.branchName}</td><td className="border px-2 py-2 font-semibold">Store Code</td><td className="border px-2 py-2"></td></tr>
            <tr><td className="border px-2 py-2 font-semibold">Location</td><td className="border px-2 py-2">{selectedBranch.locationName}</td><td className="border px-2 py-2 font-semibold">City</td><td className="border px-2 py-2">{selectedBranch.city}</td></tr>
            <tr><td className="border px-2 py-2 font-semibold">Pullout Date</td><td className="border px-2 py-2">{data.pulloutDate}</td><td className="border px-2 py-2 font-semibold">Status</td><td className="border px-2 py-2">{data.pulloutStatus}</td></tr>
            <tr><td className="border px-2 py-2 font-semibold">Closure Date</td><td className="border px-2 py-2">{data.closureDate || "Not updated"}</td><td className="border px-2 py-2 font-semibold">Destination</td><td className="border px-2 py-2">{data.destination}</td></tr>
          </tbody>
        </table>
      </div>

      <div className="store-handover-section mt-6">
        <h3 className="mb-3 text-sm font-bold text-slate-900">2. Pullout Summary</h3>
        <div className="grid gap-4 md:grid-cols-2">
          <div className="rounded-xl border border-slate-300 p-4">
            <p className="text-sm"><strong>Store Manager:</strong> {data.storeManagerName}</p>
            <p className="mt-2 text-sm"><strong>Removed By:</strong> {data.removedBy}</p>
            <p className="mt-2 text-sm"><strong>In-charge:</strong> {data.inchargeName} - {data.inchargeRole}</p>
            <p className="mt-2 text-sm"><strong>Received By:</strong> {data.receivedBy}</p>
          </div>
          <div className="rounded-xl border border-slate-300 p-4">
            <p className="text-sm"><strong>Pullout Reason:</strong> {data.pulloutReason}</p>
            <p className="mt-2 text-sm"><strong>Asset Condition:</strong> {data.assetCondition}</p>
            <p className="mt-2 text-sm"><strong>Destination:</strong> {data.destination}</p>
          </div>
        </div>
      </div>

      <div className="mt-6">
        <h3 className="mb-3 text-sm font-bold text-slate-900">3. Pulled Out Asset Details</h3>
        <table className="store-handover-table w-full border-collapse text-xs">
          <thead className="bg-slate-900 text-white"><tr><th className="border px-2 py-2 text-left">Category</th><th className="border px-2 py-2 text-left">Model</th><th className="border px-2 py-2 text-left">Item Description</th><th className="border px-2 py-2 text-left">Brand</th><th className="border px-2 py-2 text-left">Qty</th></tr></thead>
          <tbody>{assetRows.map((a) => <tr key={a.id} className="odd:bg-slate-50"><td className="border px-2 py-2 font-semibold">{a.category}</td><td className="border px-2 py-2">{a.model || "-"}</td><td className="border px-2 py-2">{a.description || a.deviceName || "-"}{a.serialNumber && <p className="mt-1 text-[10px] text-slate-500">Serial: {a.serialNumber}</p>}</td><td className="border px-2 py-2">{a.brand || "-"}</td><td className="border px-2 py-2">{a.quantity || 1}</td></tr>)}</tbody>
        </table>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">4. Pullout & Verification Confirmation</h3>
        <ul className="list-disc space-y-2 pl-8 text-sm">
          <li>All listed assets have been removed from the store/site by IT or authorized personnel.</li>
          <li>Physical count and visible asset condition were checked during pullout.</li>
          <li>Assets are handed over to the destination/receiver mentioned in this document.</li>
          <li>Any missing, damaged, or pending item should be recorded in remarks before signing.</li>
        </ul>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">5. Remarks</h3>
        <div className="grid gap-4 md:grid-cols-2">
          <div className="rounded-xl border border-slate-300 p-4">
            <p className="font-bold">IT Remarks</p>
            <p className="mt-3 text-sm leading-relaxed">{data.itRemarks || "No remarks updated."}</p>
          </div>
          <div className="rounded-xl border border-slate-300 p-4">
            <p className="font-bold">Operations / Store Remarks</p>
            <p className="mt-3 text-sm leading-relaxed">{data.operationsRemarks || "No remarks updated."}</p>
          </div>
        </div>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">6. Signatures</h3>
        <div className="grid gap-6 md:grid-cols-2">
          <div className="rounded-xl border border-slate-300 p-4"><p className="font-bold">Removed / Handed Over by IT Team</p><div className="mt-5 space-y-4 text-sm"><p>Name: ____________________________________</p><p>Designation: ______________________________</p><p>Signature: ________________________________</p><p>Date: ____ / ____ / ______</p></div></div>
          <div className="rounded-xl border border-slate-300 p-4"><p className="font-bold">Received By / Store Manager / Supervisor</p><div className="mt-5 space-y-4 text-sm"><p>Name: ____________________________________</p><p>Employee ID: ______________________________</p><p>Designation: ______________________________</p><p>Signature: ________________________________</p><p>Date: ____ / ____ / ______</p></div></div>
        </div>
      </div>

      <div className="store-handover-print-spacer hidden" />
      <div className="store-handover-flow-footer mt-8 grid grid-cols-3 items-center gap-4 px-6 py-5 text-xs font-semibold text-white" style={{ backgroundColor: footerColor }}>
        <div className="flex items-center gap-2"><span className="text-lg">☎</span><span>{footerPhone}</span></div>
        <div className="flex items-center justify-center gap-2 text-center"><span className="text-lg">📍</span><span>{footerAddress}</span></div>
        <div className="flex items-center justify-end gap-2 text-right"><span className="text-lg">◎</span><span>{footerWebsite}</span></div>
      </div>
      <div className="store-handover-print-footer hidden" style={{ backgroundColor: footerColor }}>
        <div className="flex items-center gap-2"><span>☎</span><span>{footerPhone}</span></div>
        <div className="flex items-center justify-center gap-2 text-center"><span>📍</span><span>{footerAddress}</span></div>
        <div className="flex items-center justify-end gap-2 text-right"><span>◎</span><span>{footerWebsite}</span></div>
      </div>
    </div>
    <StorePulloutPrintPages selectedBranch={selectedBranch} assets={assetRows} data={data} logo={logo} footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} />
  </div>;
}

function StoreHandoverReport({ selectedBranch, assets, data, groupBranding }) {
  const assetRows = assets.length ? assets : [];
  const declarationItems = [
    "All listed assets always remain the exclusive property of the Company.",
    "I assume full custodial accountability for safeguarding, proper usage, and operational protection of the assets within my assigned store.",
    "In the event of loss, theft, physical damage, negligence, misuse, unauthorized transfer, or improper handling, I shall bear full responsibility.",
    "The Company reserves the right to recover the replacement value, repair cost, depreciation loss, or any associated financial damage from the responsible individual(s) after internal review.",
    "Failure to report incidents immediately may result in disciplinary action, salary deduction where legally permitted, or further administrative/legal measures.",
    "This declaration is binding under Company policy and remains effective until assets are formally re-handed over to authorized personnel.",
  ];
  const branding = groupBranding || {};
  const logo = selectedBranch.logoDataUrl || branding.logoDataUrl || "";
  const footerPhone = branding.footerPhone || "Not configured";
  const footerAddress = branding.footerAddress || "Demo Head Office";
  const footerWebsite = branding.footerWebsite || "example.com";
  const footerColor = branding.footerColor || "#991b1e";

  return <div className="rounded-2xl border border-slate-200 bg-slate-100 p-4 shadow-sm">
    <div className="store-handover-screen mx-auto w-full max-w-[794px] bg-white p-8 text-slate-900 shadow-lg">
      <div className="flex items-center justify-between gap-4 border-b-4 border-slate-900 pb-5">
        <div className="flex h-20 w-52 items-center justify-center rounded-xl border border-slate-200 bg-white p-3">
          {logo ? <img src={logo} alt="Store logo" className="max-h-16 max-w-full object-contain" /> : <div className="text-center text-xs font-bold text-slate-500">STORE<br />LOGO</div>}
        </div>
        <div className="text-right text-xs leading-relaxed text-slate-600">
          <p className="font-bold text-slate-900">{selectedBranch.groupName || "Company / Client Name"}</p>
          <p>{footerAddress}</p>
          <p>{footerPhone} • {footerWebsite}</p>
        </div>
      </div>

      <div className="mt-8 text-center">
        <h2 className="text-xl font-bold uppercase tracking-wide">Store Asset Handover Form</h2>
        <p className="mt-1 text-xs text-slate-500">New Store Opening – IT / Equipment Assets</p>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">1. Store Information</h3>
        <table className="w-full border-collapse text-sm">
          <tbody>
            <tr><td className="border px-2 py-2 font-semibold">Store Name</td><td className="border px-2 py-2">{selectedBranch.branchName}</td><td className="border px-2 py-2 font-semibold">Store Code</td><td className="border px-2 py-2"></td></tr>
            <tr><td className="border px-2 py-2 font-semibold">Location</td><td className="border px-2 py-2">{selectedBranch.locationName}</td><td className="border px-2 py-2 font-semibold">City</td><td className="border px-2 py-2">{selectedBranch.city}</td></tr>
            <tr><td className="border px-2 py-2 font-semibold">Date of Handover</td><td className="border px-2 py-2">{data.handoverDate}</td><td className="border px-2 py-2 font-semibold">Status</td><td className="border px-2 py-2">{data.handoverStatus}</td></tr>
          </tbody>
        </table>
      </div>

      <div className="mt-6">
        <h3 className="mb-3 text-sm font-bold text-slate-900">2. Asset Details</h3>
        <table className="store-handover-table w-full border-collapse text-xs">
          <thead className="bg-slate-900 text-white"><tr><th className="border px-2 py-2 text-left">Category</th><th className="border px-2 py-2 text-left">Model</th><th className="border px-2 py-2 text-left">Item Description</th><th className="border px-2 py-2 text-left">Brand</th><th className="border px-2 py-2 text-left">Qty</th></tr></thead>
          <tbody>{assetRows.map((a) => <tr key={a.id} className="odd:bg-slate-50"><td className="border px-2 py-2 font-semibold">{a.category}</td><td className="border px-2 py-2">{a.model || "-"}</td><td className="border px-2 py-2">{a.description || a.deviceName || "-"}{a.serialNumber && <p className="mt-1 text-[10px] text-slate-500">Serial: {a.serialNumber}</p>}</td><td className="border px-2 py-2">{a.brand || "-"}</td><td className="border px-2 py-2">{a.quantity || 1}</td></tr>)}</tbody>
        </table>
      </div>

      <div className="store-handover-section mt-6">
        <h3 className="mb-3 text-sm font-bold text-slate-900">3. Accessories / Additional Items (If Applicable)</h3>
        <div className="grid gap-2 text-sm md:grid-cols-2"><p>☐ Power Adapters</p><p>☐ Network Cables</p><p>☐ Tablets</p><p>☐ Mouse / Keyboard</p><p>☐ SIM Card</p><p>☐ Printer Cartridges</p><p className="md:col-span-2">☐ Other: ________________________________________________</p></div>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">4. Installation & Testing Confirmation</h3>
        <ul className="list-disc space-y-2 pl-8 text-sm"><li>All listed assets have been installed, configured, and tested by IT.</li><li>Devices are fully operational at the time of handover.</li><li>Store representatives have verified physical count, serial numbers, and working condition.</li></ul>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">5. Acknowledgement & Legal Responsibility Declaration</h3>
        <p className="text-sm leading-relaxed">I, the undersigned Store Manager / Supervisor, hereby acknowledge receipt of the above-mentioned company-owned assets in good working condition. I expressly agree and legally undertake the following obligations:</p>
        <ol className="mt-4 list-decimal space-y-2 pl-6 text-sm leading-relaxed">{declarationItems.map((item) => <li key={item}>{item}</li>)}</ol>
      </div>

      <div className="store-handover-section mt-8">
        <h3 className="mb-3 text-sm font-bold text-slate-900">6. Signatures</h3>
        <div className="grid gap-6 md:grid-cols-2">
          <div className="rounded-xl border border-slate-300 p-4"><p className="font-bold">Handed Over by IT Team</p><div className="mt-5 space-y-4 text-sm"><p>Name: ____________________________________</p><p>Designation: ______________________________</p><p>Signature: ________________________________</p><p>Date: ____ / ____ / ______</p></div></div>
          <div className="rounded-xl border border-slate-300 p-4"><p className="font-bold">Received By (Store Manager / Supervisor)</p><div className="mt-5 space-y-4 text-sm"><p>Name: ____________________________________</p><p>Employee ID: ______________________________</p><p>Designation: ______________________________</p><p>Signature: ________________________________</p><p>Date: ____ / ____ / ______</p></div></div>
        </div>
      </div>

      <div className="store-handover-print-spacer hidden" />
      <div className="store-handover-flow-footer mt-8 grid grid-cols-3 items-center gap-4 px-6 py-5 text-xs font-semibold text-white" style={{ backgroundColor: footerColor }}>
        <div className="flex items-center gap-2"><span className="text-lg">☎</span><span>{footerPhone}</span></div>
        <div className="flex items-center justify-center gap-2 text-center"><span className="text-lg">📍</span><span>{footerAddress}</span></div>
        <div className="flex items-center justify-end gap-2 text-right"><span className="text-lg">◎</span><span>{footerWebsite}</span></div>
      </div>
      <div className="store-handover-print-footer hidden" style={{ backgroundColor: footerColor }}>
        <div className="flex items-center gap-2"><span>☎</span><span>{footerPhone}</span></div>
        <div className="flex items-center justify-center gap-2 text-center"><span>📍</span><span>{footerAddress}</span></div>
        <div className="flex items-center justify-end gap-2 text-right"><span>◎</span><span>{footerWebsite}</span></div>
      </div>
    </div>
    <StoreHandoverPrintPages selectedBranch={selectedBranch} assets={assetRows} data={data} logo={logo} footerPhone={footerPhone} footerAddress={footerAddress} footerWebsite={footerWebsite} footerColor={footerColor} declarationItems={declarationItems} />
  </div>;
}

function Report({ type, selectedBranch, assets, network, data }) {
  const isPullout = type === "pullout";
  return <div className="rounded-2xl border border-slate-200 bg-slate-100 p-4 shadow-sm">
    <div className="printable-report mx-auto w-full max-w-[794px] rounded-xl bg-white p-8 text-slate-900 shadow-lg">
      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
        <div><h3 className="text-xl font-bold uppercase tracking-wide">{isPullout ? "IT Asset Pullout Report" : "Department Handover Report"}</h3></div>
        <Badge type={(data.pulloutStatus || data.handoverStatus) === "Signed / Completed" ? "success" : "warning"}>{data.pulloutStatus || data.handoverStatus}</Badge>
      </div>
      <div className="pdf-row mt-5 grid gap-3 md:grid-cols-3"><div className="pdf-card rounded-xl border border-slate-200 p-4"><p className="text-xs uppercase text-slate-500">Branch</p><p className="font-bold">{selectedBranch.branchName}</p></div><div className="pdf-card rounded-xl border border-slate-200 p-4"><p className="text-xs uppercase text-slate-500">Location</p><p className="font-bold">{selectedBranch.locationName}, {selectedBranch.city}</p></div><div className="pdf-card rounded-xl border border-slate-200 p-4"><p className="text-xs uppercase text-slate-500">Date</p><p className="font-bold">{data.pulloutDate || data.handoverDate}</p></div></div>
      <div className="pdf-two pdf-section mt-6 grid gap-4 md:grid-cols-2"><div className="pdf-card rounded-2xl border border-slate-200 p-4"><h4 className="font-bold">People & Approval</h4><div className="mt-3 space-y-2 text-sm"><p><strong>Store Manager:</strong> {data.storeManagerName}</p><p><strong>{isPullout ? "Removed By" : "Completed By"}:</strong> {data.removedBy || data.completedPersons}</p><p><strong>In-charge:</strong> {data.inchargeName} - {data.inchargeRole}</p><p><strong>Received By:</strong> {data.receivedBy}</p></div></div><div className="pdf-card rounded-2xl border border-slate-200 p-4"><h4 className="font-bold">{isPullout ? "Pullout Summary" : "Network Summary"}</h4><div className="mt-3 space-y-2 text-sm">{isPullout ? <><p><strong>Closure Date:</strong> {data.closureDate || "Not updated"}</p><p><strong>Reason:</strong> {data.pulloutReason}</p><p><strong>Destination:</strong> {data.destination}</p><p><strong>Condition:</strong> {data.assetCondition}</p></> : <><p><strong>ISP:</strong> {network?.isp || "To be updated"}</p><p><strong>Router IP:</strong> {network?.routerIp || "To be updated"}</p><p><strong>NVR IP:</strong> {network?.nvrIp || "To be updated"}</p><p><strong>POS Range:</strong> {network?.posRange || "To be updated"}</p></>}</div></div></div>
      <div className="pdf-section mt-5"><h4 className="font-bold">{isPullout ? "Pulled Out Devices / Assets" : "Installed Devices / Assets"}</h4><div className="mt-3 overflow-hidden rounded-2xl border border-slate-200"><table className="pdf-table w-full text-left text-sm"><thead className="bg-slate-100 text-slate-600"><tr><th className="px-3 py-2">#</th><th className="px-3 py-2">Category</th><th className="px-3 py-2">Device</th><th className="px-3 py-2">Serial No.</th><th className="px-3 py-2">IP Address</th><th className="px-3 py-2">Installed Location</th><th className="px-3 py-2">Status</th></tr></thead><tbody className="divide-y divide-slate-200">{assets.map((a, i) => <tr key={a.id}><td className="px-3 py-2">{i + 1}</td><td className="px-3 py-2">{a.category}</td><td className="px-3 py-2 font-semibold">{a.deviceName}</td><td className="px-3 py-2">{a.serialNumber}</td><td className="px-3 py-2">{a.ipAddress}</td><td className="px-3 py-2">{a.installedLocation}</td><td className="px-3 py-2">{a.status}</td></tr>)}</tbody></table></div></div>
      <div className="pdf-two pdf-section mt-6 grid gap-4 md:grid-cols-2"><div className="pdf-card rounded-2xl border border-slate-200 p-4"><h4 className="font-bold">IT Remarks</h4><p className="mt-2 text-sm text-slate-600">{data.itRemarks}</p></div><div className="pdf-card rounded-2xl border border-slate-200 p-4"><h4 className="font-bold">Operations Remarks</h4><p className="mt-2 text-sm text-slate-600">{data.operationsRemarks}</p></div></div>
      <div className="pdf-signatures pdf-section mt-8 grid gap-4 md:grid-cols-3"><div className="pdf-card rounded-2xl border border-slate-300 p-4 text-center"><p className="text-sm text-slate-500">{isPullout ? "Removed By" : "Prepared / Completed By"}</p><div className="mt-8 border-t border-slate-300 pt-2 font-semibold">{data.removedBy || data.completedPersons}</div></div><div className="pdf-card rounded-2xl border border-slate-300 p-4 text-center"><p className="text-sm text-slate-500">Checked By / In-charge</p><div className="mt-8 border-t border-slate-300 pt-2 font-semibold">{data.inchargeName}</div></div><div className="pdf-card rounded-2xl border border-slate-300 p-4 text-center"><p className="text-sm text-slate-500">Received By / Manager</p><div className="mt-8 border-t border-slate-300 pt-2 font-semibold">{data.storeManagerName}</div></div></div>
    </div>
  </div>;
}

function SettingsSection({ theme, setTheme, isAdmin, users, forms, updateForm, setUsers, currentUser, setCurrentUser, deleteUser, exportDatabase, restoreInputRef, restoreDatabaseFile, resetData, syncFromServer, serverMessage, cleanImportedExcelRemarks, assetProducts = [], assetCategories = [], setAssetCategories, assets = [], openModal, deleteProduct, resetRequests = [], setResetRequests }) {
  const [resetTarget, setResetTarget] = useState(null);
  const [resetPassword, setResetPassword] = useState({ password: "", confirm: "" });
  const [resetError, setResetError] = useState("");
  function createUser(e) {
    e.preventDefault();
    const f = forms.newUser;
    if (!f.displayName || !f.username || !f.password) return alert("Complete display name, username and password.");
    if (users.some((u) => u.username.toLowerCase() === f.username.toLowerCase())) return alert("Username already exists.");
    setUsers((x) => [{ id: makeId(), displayName: f.displayName, username: f.username, password: f.password, role: f.role }, ...x]);
    updateForm("newUser", "displayName", ""); updateForm("newUser", "username", ""); updateForm("newUser", "password", "");
  }
  function resetPass(e) {
    e.preventDefault();
    const f = forms.adminPassword;
    const u = users.find((x) => x.id === currentUser.id);
    if (u.password !== f.currentPassword) return alert("Current password is incorrect.");
    if (!f.newPassword || f.newPassword !== f.confirmPassword) return alert("New passwords do not match.");
    setUsers((x) => x.map((v) => v.id === u.id ? { ...v, password: f.newPassword } : v));
    setCurrentUser((x) => ({ ...x, password: f.newPassword }));
    alert("Password updated.");
  }
  function openResetRequest(request) {
    const user = users.find((item) => item.username.toLowerCase() === request.username.toLowerCase());
    if (!user) {
      setResetRequests((items) => items.map((item) => item.id === request.id ? { ...item, status: "Reviewed", resolvedAt: new Date().toISOString(), resolvedBy: currentUser.displayName } : item));
      return;
    }
    setResetTarget({ request, user });
    setResetPassword({ password: "", confirm: "" });
    setResetError("");
  }
  function completeResetRequest(e) {
    e.preventDefault();
    if (!resetTarget) return;
    if (resetPassword.password.length < 8) return setResetError("Temporary password must contain at least 8 characters.");
    if (resetPassword.password !== resetPassword.confirm) return setResetError("Passwords do not match.");
    setUsers((items) => items.map((item) => item.id === resetTarget.user.id ? { ...item, password: resetPassword.password } : item));
    setResetRequests((items) => items.map((item) => item.id === resetTarget.request.id ? { ...item, status: "Resolved", resolvedAt: new Date().toISOString(), resolvedBy: currentUser.displayName } : item));
    setResetTarget(null);
    setResetPassword({ password: "", confirm: "" });
    setResetError("");
  }
  function removeResetRequest(id) {
    setResetRequests((items) => items.filter((item) => item.id !== id));
  }
  const [newCategory, setNewCategory] = useState({ name: "", description: "" });
  function createCategory(e) {
    e.preventDefault();
    const name = newCategory.name.trim();
    if (!name) return alert("Enter a category name.");
    if (assetCategories.some((item) => normalizeExcelText(item.name) === normalizeExcelText(name))) return alert("This category already exists.");
    setAssetCategories((items) => [...items, { id: makeId(), name, description: newCategory.description.trim(), isActive: true }].sort((a, b) => a.name.localeCompare(b.name)));
    setNewCategory({ name: "", description: "" });
  }
  function toggleCategory(category) {
    const used = assetProducts.some((product) => normalizeExcelText(product.category) === normalizeExcelText(category.name)) || assets.some((asset) => normalizeExcelText(asset.category) === normalizeExcelText(category.name));
    if (category.isActive !== false && used) {
      setAssetCategories((items) => items.map((item) => item.id === category.id ? { ...item, isActive: false } : item));
      return;
    }
    setAssetCategories((items) => items.map((item) => item.id === category.id ? { ...item, isActive: item.isActive === false } : item));
  }
  return <section className="mt-6 space-y-5">
    <div className="rounded-3xl border border-slate-200 bg-white p-6 shadow-sm">
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div><h3 className="text-2xl font-bold">Settings</h3><p className="mt-1 text-sm text-slate-500">Professional control center for users, templates, product master, backup and cleanup tools.</p></div>
        <Badge type="info">{serverMessage}</Badge>
      </div>
    </div>

    {isAdmin && <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <h4 className="text-lg font-bold">Template Tools</h4><p className="mt-1 text-sm text-slate-500">Branch, location and IT asset template downloads. Asset Product List and Standard Branch Asset Templates are now inside the IT Assets page.</p>
      <div className="mt-4 grid gap-2 md:grid-cols-4"><button onClick={() => openModal("template", { type: "branch" })} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Branch Template</button><button onClick={() => openModal("template", { type: "location" })} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Location Template</button><button onClick={() => openModal("template", { type: "asset" })} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-700">IT Asset Template</button><button onClick={() => openModal("template", { type: "branchAssetTemplate" })} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Branch Asset Template</button></div>
    </div>}

    {isAdmin && <div className="grid gap-4 xl:grid-cols-3">
      <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><h4 className="text-lg font-bold">Database / Cloud Sync</h4><p className="mt-1 text-sm text-slate-500">Backup, restore and sync Firebase data.</p><div className="mt-4 grid gap-2"><button onClick={() => syncFromServer(true)} className="rounded-2xl border border-blue-200 bg-blue-50 px-4 py-2 text-sm font-semibold text-blue-700">Sync from Firebase</button><button onClick={exportDatabase} className="rounded-2xl border border-green-200 bg-green-50 px-4 py-2 text-sm font-semibold text-green-700">Export Database Backup</button><button onClick={() => restoreInputRef.current?.click()} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Import / Restore Backup</button><input ref={restoreInputRef} type="file" accept="application/json,.json" onChange={restoreDatabaseFile} className="hidden" /></div></div>
      <div className="rounded-3xl border border-amber-200 bg-amber-50 p-5 shadow-sm"><h4 className="text-lg font-bold text-amber-800">Cleanup Tools</h4><p className="mt-1 text-sm text-amber-700">Clean imported placeholder text from existing records.</p><button onClick={cleanImportedExcelRemarks} className="mt-4 rounded-2xl border border-amber-300 bg-white px-4 py-2 text-sm font-semibold text-amber-700">Clean Imported Excel Remarks</button></div>
      <div className="rounded-3xl border border-red-200 bg-red-50 p-5 shadow-sm"><h4 className="text-lg font-bold text-red-800">Danger Zone</h4><p className="mt-1 text-sm text-red-700">Reset is separated here to avoid accidental clicks. Export backup before reset.</p><button onClick={resetData} className="mt-4 rounded-2xl border border-red-300 bg-white px-4 py-2 text-sm font-semibold text-red-700">Reset Application Data</button></div>
    </div>}

    {isAdmin && <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between"><div><h4 className="text-lg font-bold">Asset Category Master</h4><p className="mt-1 text-sm text-slate-500">Add categories once and reuse them in Product List, Asset Register and templates.</p></div><Badge type="info">{assetCategories.filter((item) => item.isActive !== false).length} Active</Badge></div>
      <form onSubmit={createCategory} className="mt-4 grid gap-3 md:grid-cols-[1fr_1.5fr_auto]"><TextInput value={newCategory.name} onChange={(value) => setNewCategory((current) => ({ ...current, name: value }))} placeholder="New category name" /><TextInput value={newCategory.description} onChange={(value) => setNewCategory((current) => ({ ...current, description: value }))} placeholder="Description (optional)" /><button className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">Add Category</button></form>
      <div className="mt-4 grid gap-2 sm:grid-cols-2 xl:grid-cols-3">{assetCategories.map((category) => { const usage = assetProducts.filter((product) => normalizeExcelText(product.category) === normalizeExcelText(category.name)).length + assets.filter((asset) => normalizeExcelText(asset.category) === normalizeExcelText(category.name)).length; return <div key={category.id} className="flex items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-3"><div className="min-w-0"><p className="truncate text-sm font-semibold text-slate-900">{category.name}</p><p className="text-xs text-slate-500">{usage} linked record(s) • {category.isActive === false ? "Inactive" : "Active"}</p></div><button type="button" onClick={() => toggleCategory(category)} className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${category.isActive === false ? "border-emerald-200 bg-emerald-50 text-emerald-700" : "border-amber-200 bg-amber-50 text-amber-700"}`}>{category.isActive === false ? "Activate" : "Deactivate"}</button></div>; })}</div>
    </div>}

    <div className="grid gap-5 xl:grid-cols-2">
      <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><h4 className="text-lg font-bold">Appearance</h4><div className="mt-4 flex rounded-2xl bg-slate-100 p-1"><button onClick={() => setTheme("light")} className={`flex-1 rounded-xl px-4 py-2 text-sm font-semibold ${theme === "light" ? "bg-white text-slate-900 shadow-sm" : "text-slate-600"}`}>Light Mode</button><button onClick={() => setTheme("dark")} className={`flex-1 rounded-xl px-4 py-2 text-sm font-semibold ${theme === "dark" ? "bg-slate-900 text-white shadow-sm" : "text-slate-600"}`}>Dark Mode</button></div></div>
      <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><h4 className="text-lg font-bold">Change My Password</h4><p className="mt-1 text-sm text-slate-500">Use this after signing in with a temporary password.</p><form onSubmit={resetPass} className="mt-4 grid gap-3"><Field label="Current Password" required><TextInput type="password" value={forms.adminPassword.currentPassword} onChange={(v) => updateForm("adminPassword", "currentPassword", v)} /></Field><Field label="New Password" required><TextInput type="password" value={forms.adminPassword.newPassword} onChange={(v) => updateForm("adminPassword", "newPassword", v)} /></Field><Field label="Confirm New Password" required><TextInput type="password" value={forms.adminPassword.confirmPassword} onChange={(v) => updateForm("adminPassword", "confirmPassword", v)} /></Field><button className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">Update My Password</button></form></div>
    </div>

    {isAdmin && <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div><h4 className="text-lg font-bold">Password Reset Requests</h4><p className="mt-1 text-sm text-slate-500">Review requests submitted from the login page and issue a temporary password.</p></div><Badge type={resetRequests.some((item) => item.status === "Pending" || item.status === "Needs Verification") ? "warning" : "success"}>{resetRequests.filter((item) => item.status === "Pending" || item.status === "Needs Verification").length} Pending</Badge></div>
      <div className="mt-5 space-y-3">{resetRequests.length === 0 ? <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-5 text-sm text-slate-500">No password reset requests.</div> : resetRequests.map((request) => { const matched = users.some((user) => user.username.toLowerCase() === request.username.toLowerCase()); const open = request.status === "Pending" || request.status === "Needs Verification"; return <div key={request.id} className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between"><div className="min-w-0"><div className="flex flex-wrap items-center gap-2"><p className="font-semibold text-slate-900">{request.displayName || request.username}</p><Badge type={request.status === "Resolved" ? "success" : request.status === "Needs Verification" ? "danger" : request.status === "Reviewed" ? "info" : "warning"}>{request.status}</Badge></div><p className="mt-1 text-xs text-slate-500">Username: {request.username} • Contact: {request.contact}</p><p className="mt-1 text-xs text-slate-500">Requested: {new Date(request.requestedAt).toLocaleString()}</p>{request.message && <p className="mt-2 text-sm text-slate-600">{request.message}</p>}{request.resolvedAt && <p className="mt-2 text-xs text-slate-500">Handled by {request.resolvedBy || "Admin"} on {new Date(request.resolvedAt).toLocaleString()}</p>}</div><div className="flex shrink-0 flex-wrap gap-2">{open && <button type="button" onClick={() => openResetRequest(request)} className="rounded-xl bg-slate-900 px-4 py-2 text-xs font-semibold text-white">{matched ? "Reset Password" : "Mark Reviewed"}</button>}<button type="button" onClick={() => removeResetRequest(request.id)} className="rounded-xl border border-red-200 bg-white px-4 py-2 text-xs font-semibold text-red-700">Delete</button></div></div></div>; })}</div>
    </div>}

    {isAdmin && <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><div className="flex items-center justify-between"><div><h4 className="text-lg font-bold">User Creation & Access Level</h4><p className="mt-1 text-sm text-slate-500">Create Admin or User accounts.</p></div><Badge type="info">{users.length} Users</Badge></div><form onSubmit={createUser} className="mt-5 grid gap-4 md:grid-cols-2 xl:grid-cols-4"><Field label="Display Name" required><TextInput value={forms.newUser.displayName} onChange={(v) => updateForm("newUser", "displayName", v)} /></Field><Field label="Username" required><TextInput value={forms.newUser.username} onChange={(v) => updateForm("newUser", "username", v)} /></Field><Field label="Password" required><TextInput type="password" value={forms.newUser.password} onChange={(v) => updateForm("newUser", "password", v)} /></Field><Field label="Access Level" required><SelectInput value={forms.newUser.role} onChange={(v) => updateForm("newUser", "role", v)}><option value="user">User - View Only</option><option value="admin">Admin - Full Access</option></SelectInput></Field><button className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white md:col-span-2 xl:col-span-4">Create User</button></form><div className="mt-6"><Table headers={["Display Name", "Username", "Access Level", "Actions"]} rows={users} renderDesktop={(u) => <tr key={u.id} className="hover:bg-slate-50"><td className="px-4 py-3 font-semibold">{u.displayName}</td><td className="px-4 py-3 text-slate-600">{u.username}</td><td className="px-4 py-3"><Badge type={u.role === "admin" ? "success" : "info"}>{u.role}</Badge></td><td className="px-4 py-3"><button onClick={() => deleteUser(u.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button></td></tr>} renderMobile={(u) => <MobileCard key={u.id} kicker="User" title={u.displayName} subtitle={u.username} status={<Badge type={u.role === "admin" ? "success" : "info"}>{u.role}</Badge>} actions={<button onClick={() => deleteUser(u.id)} className="rounded-lg border border-red-200 px-2 py-1 text-xs font-semibold text-red-700">Delete</button>} />} /></div></div>}
    {resetTarget && <Modal title="Reset User Password" subtitle={`Set a temporary password for ${resetTarget.user.displayName} (${resetTarget.user.username}).`} onClose={() => setResetTarget(null)} max="max-w-lg">{resetError && <ErrorBox text={resetError} />}<form onSubmit={completeResetRequest} className="mt-6 space-y-4"><Field label="Temporary Password" required><TextInput type="password" value={resetPassword.password} onChange={(value) => setResetPassword((current) => ({ ...current, password: value }))} /></Field><Field label="Confirm Password" required><TextInput type="password" value={resetPassword.confirm} onChange={(value) => setResetPassword((current) => ({ ...current, confirm: value }))} /></Field><p className="rounded-2xl bg-amber-50 p-4 text-xs leading-5 text-amber-800">Share the temporary password securely and ask the user to change it after signing in.</p><div className="flex gap-2"><button type="button" onClick={() => setResetTarget(null)} className="flex-1 rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button><button type="submit" className="flex-1 rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">Reset Password</button></div></form></Modal>}
  </section>;
}

function ModalRouter({ modal, close, forms, updateForm, formError, submit, groups, selectedBranch, selectedBranchAssets, locations, branches, selectedNetwork, editId, openModal, excelImport, excelActions, branchImport, branchActions, locationImport, locationActions, productImport, productActions, templateBuilder, templateActions, assetProducts, assetCategories, addAssetCategory, assetTemplates, applyBranchAssetTemplate, vaultUnlocked, allAssets, credentials }) {
  if (modal === "location") return <Modal title={editId ? "Edit Location" : "Add New Location"} subtitle={editId ? "Update mall/site details. Branch records using this location will be updated safely." : "Add mall or site location once."} onClose={close} max="max-w-2xl">{formError && <ErrorBox text={formError} />}<form onSubmit={submit.location} className="mt-6 grid gap-4 md:grid-cols-2"><Field label="Location / Mall Name" required><TextInput value={forms.location.name} onChange={(v) => updateForm("location", "name", v)} /></Field><Field label="City" required><TextInput value={forms.location.city} onChange={(v) => updateForm("location", "city", v)} /></Field><Field label="Emirate" required><SelectInput value={forms.location.emirate} onChange={(v) => updateForm("location", "emirate", v)}>{["Dubai", "Abu Dhabi", "Sharjah", "Ajman", "Ras Al Khaimah", "Fujairah", "Umm Al Quwain"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Location Type" required><SelectInput value={forms.location.type} onChange={(v) => updateForm("location", "type", v)}>{["Mall", "Street", "Airport", "Office", "Warehouse", "Cloud Kitchen", "Other"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Notes</span><TextArea value={forms.location.notes || ""} onChange={(v) => updateForm("location", "notes", v)} /></label><ModalButtons save={editId ? "Update Location" : "Save Location"} close={close} /></form></Modal>;
  if (modal === "branch") { const group = groups.find((g) => g.name === forms.branch.groupName) || groups[0] || defaultGroups[0]; return <Modal title={editId ? "Edit Branch Details" : "Add New Branch"} subtitle="Create branch first. Then add branch assets and network setup inside branch details." onClose={close} max="max-w-3xl">{formError && <ErrorBox text={formError} />}<form onSubmit={submit.branch} className="mt-6 grid gap-4 md:grid-cols-2"><Field label="Main Group" required><SelectInput value={forms.branch.groupName} onChange={(v) => { updateForm("branch", "groupName", v); const g = groups.find((x) => x.name === v); updateForm("branch", "conceptName", g?.concepts?.[0] || ""); }}>{groups.map((g) => <option key={g.id}>{g.name}</option>)}</SelectInput></Field><Field label="Concept / Restaurant" required><SelectInput value={forms.branch.conceptName} onChange={(v) => updateForm("branch", "conceptName", v)}>{group.concepts.map((c) => <option key={c}>{c}</option>)}</SelectInput></Field><Field label="Branch Display Name" required><TextInput value={forms.branch.branchName} onChange={(v) => updateForm("branch", "branchName", v)} /></Field><Field label="Location / Mall" required><SelectInput value={forms.branch.locationName} onChange={(v) => updateForm("branch", "locationName", v)}>{locations.map((l) => <option key={l.id} value={l.name}>{l.name} - {l.city}</option>)}</SelectInput></Field><Field label="Branch Type" required><SelectInput value={forms.branch.branchType} onChange={(v) => updateForm("branch", "branchType", v)}>{["QSR", "Dining", "Cloud Kitchen", "Kiosk", "Office", "Warehouse"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Opening Date" required><TextInput type="date" value={forms.branch.openingDate} onChange={(v) => updateForm("branch", "openingDate", v)} /></Field><Field label="IT Setup Completed Date" required><TextInput type="date" value={forms.branch.completedDate} onChange={(v) => updateForm("branch", "completedDate", v)} /></Field><Field label="Current Status" required><SelectInput value={forms.branch.status} onChange={(v) => updateForm("branch", "status", v)}>{["Planning", "Installation Started", "Pending Handover", "Opening Today", "Live", "Closed"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Branch Manager"><TextInput value={forms.branch.branchManager} onChange={(v) => updateForm("branch", "branchManager", v)} /></Field><Field label="Contact Number"><TextInput value={forms.branch.contactNumber} onChange={(v) => updateForm("branch", "contactNumber", v)} /></Field><Field label="Prepared By"><TextInput value={forms.branch.preparedBy} onChange={(v) => updateForm("branch", "preparedBy", v)} /></Field><label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Notes</span><TextArea value={forms.branch.notes} onChange={(v) => updateForm("branch", "notes", v)} /></label><div className="flex flex-wrap gap-2 md:col-span-2"><button type="submit" className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">{editId ? "Save Branch Changes" : "Submit & Save Branch"}</button><button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button><button type="button" onClick={() => openModal("location")} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Add Missing Location</button></div></form></Modal>; }
  if (modal === "groupBranding") return <GroupBrandingModal forms={forms} updateForm={updateForm} submit={submit.groupBranding} formError={formError} close={close} />;
  if (modal === "template") return <TemplateBuilderModal templateBuilder={templateBuilder} actions={templateActions} close={close} />;
  if (modal === "assetImport") return <ExcelAssetImportModal excelImport={excelImport} actions={excelActions} selectedBranch={selectedBranch} close={close} />;
  if (modal === "branchImport") return <ExcelBranchImportModal branchImport={branchImport} actions={branchActions} close={close} />;
  if (modal === "locationImport") return <ExcelLocationImportModal locationImport={locationImport} actions={locationActions} close={close} />;
  if (modal === "productImport") return <ExcelProductImportModal productImport={productImport} actions={productActions} close={close} />;
  if (modal === "asset") return <AssetModal forms={forms} updateForm={updateForm} submit={submit.asset} formError={formError} close={close} editId={editId} assetProducts={assetProducts} assetCategories={assetCategories} />;
  if (modal === "product") return <AssetProductModal forms={forms} updateForm={updateForm} submit={submit.product} formError={formError} close={close} editId={editId} assetCategories={assetCategories} addAssetCategory={addAssetCategory} />;
  if (modal === "assetTemplate") return <AssetTemplateModal forms={forms} updateForm={updateForm} submit={submit.assetTemplate} formError={formError} close={close} editId={editId} assetProducts={assetProducts} />;
  if (modal === "applyAssetTemplate") return <ApplyAssetTemplateModal close={close} selectedBranch={selectedBranch} assetTemplates={assetTemplates} assetProducts={assetProducts} applyBranchAssetTemplate={applyBranchAssetTemplate} />;
  if (modal === "network") return <NetworkModal forms={forms} updateForm={updateForm} submit={submit.network} formError={formError} close={close} selectedBranch={selectedBranch} />;
  if (modal === "fault") return <FaultModal forms={forms} updateForm={updateForm} submit={submit.fault} formError={formError} close={close} branches={branches} editId={editId} />;
  if (modal === "handover") return <HandoverModal type="handover" forms={forms} updateForm={updateForm} submit={submit.handover} formError={formError} close={close} assets={selectedBranchAssets} selectedBranch={selectedBranch} />;
  if (modal === "pullout") return <HandoverModal type="pullout" forms={forms} updateForm={updateForm} submit={submit.pullout} formError={formError} close={close} assets={selectedBranchAssets} selectedBranch={selectedBranch} />;
  if (modal === "profile") return <ProfileModal forms={forms} updateForm={updateForm} submit={submit.profile} formError={formError} close={close} editId={editId} />;
  if (modal === "credential") return <CredentialModal forms={forms} updateForm={updateForm} submit={submit.credential} formError={formError} close={close} editId={editId} branches={branches} assets={allAssets || selectedBranchAssets || []} credentials={credentials || []} vaultUnlocked={vaultUnlocked} />;
  if (modal === "cctv") return <CctvModal forms={forms} updateForm={updateForm} submit={submit.cctv} formError={formError} close={close} />;
  return null;
}
function ErrorBox({ text }) { return <div className="mt-4 rounded-2xl bg-red-50 p-4 text-sm font-semibold text-red-700">{text}</div>; }
function ModalButtons({ save, close }) { return <div className="app-modal-actions mt-4 flex flex-wrap gap-2 pt-4 md:col-span-2 xl:col-span-3"><button type="submit" className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">{save}</button><button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button></div>; }
function CredentialModal({ forms, updateForm, submit, formError, close, editId, branches, assets, credentials, vaultUnlocked }) {
  const f = normalizeCredentialRecord(forms.credential);
  const sharedAccounts = credentials.map(normalizeCredentialRecord).filter((item) => item.type === "shared");
  const branchAssets = assets.filter((asset) => !f.branchId || asset.branchId === Number(f.branchId));
  const selectedAsset = assets.find((asset) => asset.id === Number(f.linkedAssetId));
  const selectedShared = sharedAccounts.find((item) => item.id === Number(f.linkedCredentialId));
  const changeType = (value) => {
    const base = { ...emptyCredential, type: value };
    if (value === "shared") base.scopeType = "QSR";
    if (value === "access") base.status = "Active";
    updateForm("credential", "type", value);
    Object.entries(base).forEach(([key, fieldValue]) => { if (key !== "type") updateForm("credential", key, fieldValue); });
  };
  const selectAsset = (value) => {
    updateForm("credential", "linkedAssetId", value);
    const asset = assets.find((entry) => entry.id === Number(value));
    if (asset) {
      updateForm("credential", "deviceName", f.deviceName || asset.deviceName || asset.category || "");
      updateForm("credential", "category", asset.category || f.category || "Other");
      updateForm("credential", "ipAddress", f.ipAddress || asset.ipAddress || "");
      if (!f.branchId && asset.branchId) updateForm("credential", "branchId", String(asset.branchId));
    }
  };
  if (!vaultUnlocked) return <Modal title="Credential Vault Locked" onClose={close}><p className="mt-5 text-sm text-slate-600">Unlock the Credential Vault before adding or editing credentials.</p></Modal>;
  return <Modal title={editId ? "Edit Vault Record" : "Add Vault Record"} subtitle="Passwords are encrypted before the record is saved." onClose={close} max="max-w-6xl">
    {formError && <ErrorBox text={formError} />}
    <form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
      <Field label="Record Type" required><SelectInput value={f.type} onChange={changeType}><option value="shared">Shared Account</option><option value="device">Device Credential</option><option value="access">User Access Assignment</option></SelectInput></Field>
      <Field label="Status"><SelectInput value={f.status} onChange={(value) => updateForm("credential", "status", value)}>{["Active", "Inactive", "Pending", "Revoked", "Expired"].map((item) => <option key={item}>{item}</option>)}</SelectInput></Field>
      <div className="hidden xl:block" />

      {f.type === "shared" && <>
        <Field label="Platform / System" required><TextInput value={f.platform} onChange={(value) => updateForm("credential", "platform", value)} placeholder="Hik-Connect / No-IP / Portal" /></Field>
        <Field label="Account Scope" required><SelectInput value={f.scopeType} onChange={(value) => updateForm("credential", "scopeType", value)}>{credentialScopes.map((item) => <option key={item}>{item}</option>)}</SelectInput></Field>
        <Field label="Scope Name"><TextInput value={f.scopeName} onChange={(value) => updateForm("credential", "scopeName", value)} placeholder="QSR Supervisors / DTF / Management" /></Field>
        <Field label="Account Name" required><TextInput value={f.accountName} onChange={(value) => updateForm("credential", "accountName", value)} placeholder="Hik-Connect - QSR Supervisors" /></Field>
        <Field label="Login Email"><TextInput type="email" value={f.email} onChange={(value) => updateForm("credential", "email", value)} /></Field>
        <Field label="Username"><TextInput value={f.username} onChange={(value) => updateForm("credential", "username", value)} /></Field>
        <Field label={editId ? "New Password" : "Current Password"}><PasswordInput value={f.password} onChange={(value) => updateForm("credential", "password", value)} placeholder={editId ? "Leave blank to keep existing password" : "Enter current password"} /></Field>
        <Field label="Account Owner"><TextInput value={f.owner} onChange={(value) => updateForm("credential", "owner", value)} placeholder="IT Department" /></Field>
        <Field label="MFA Status"><SelectInput value={f.mfaStatus} onChange={(value) => updateForm("credential", "mfaStatus", value)}>{["Enabled", "Not Enabled", "SMS", "Authenticator", "Email"].map((item) => <option key={item}>{item}</option>)}</SelectInput></Field>
        <Field label="Recovery Email"><TextInput type="email" value={f.recoveryEmail} onChange={(value) => updateForm("credential", "recoveryEmail", value)} /></Field>
        <Field label="Recovery Mobile"><TextInput value={f.mobile} onChange={(value) => updateForm("credential", "mobile", value)} /></Field>
        <Field label="Last Password Changed"><TextInput type="date" value={f.lastPasswordChange} onChange={(value) => updateForm("credential", "lastPasswordChange", value)} /></Field>
        <Field label="Next Review Date"><TextInput type="date" value={f.nextPasswordChange} onChange={(value) => updateForm("credential", "nextPasswordChange", value)} /></Field>
      </>}

      {f.type === "device" && <>
        <Field label="Branch / Location" required><SelectInput value={f.branchId} onChange={(value) => { updateForm("credential", "branchId", value); updateForm("credential", "linkedAssetId", ""); }}><option value="">Select branch</option>{branches.map((branch) => <option key={branch.id} value={branch.id}>{branch.branchName} - {branch.locationName}</option>)}</SelectInput></Field>
        <Field label="Device Type" required><SelectInput value={f.category} onChange={(value) => updateForm("credential", "category", value)}>{["NVR", "DVR", "Server", "NAS", "Router / Firewall", "Network Switch", "Access Point", "Biometric", "POS", "Other"].map((item) => <option key={item}>{item}</option>)}</SelectInput></Field>
        <Field label="Linked Asset"><SelectInput value={f.linkedAssetId} onChange={selectAsset}><option value="">Not linked</option>{branchAssets.map((asset) => <option key={asset.id} value={asset.id}>{asset.deviceName} - {asset.serialNumber || "No serial"}</option>)}</SelectInput></Field>
        <Field label="Device Name" required><TextInput value={f.deviceName} onChange={(value) => updateForm("credential", "deviceName", value)} placeholder="Main CCTV NVR" /></Field>
        <Field label="IP Address"><TextInput value={f.ipAddress} onChange={(value) => updateForm("credential", "ipAddress", value)} placeholder="192.168.20.10" /></Field>
        <Field label="URL / Hostname"><TextInput value={f.url} onChange={(value) => updateForm("credential", "url", value)} placeholder="https://... or hostname" /></Field>
        <Field label="Port"><TextInput value={f.port} onChange={(value) => updateForm("credential", "port", value)} placeholder="8000" /></Field>
        <Field label="Username"><TextInput value={f.username} onChange={(value) => updateForm("credential", "username", value)} /></Field>
        <Field label={editId ? "New Password" : "Password"}><PasswordInput value={f.password} onChange={(value) => updateForm("credential", "password", value)} placeholder={editId ? "Leave blank to keep existing password" : "Enter password"} /></Field>
        <Field label="Last Password Changed"><TextInput type="date" value={f.lastPasswordChange} onChange={(value) => updateForm("credential", "lastPasswordChange", value)} /></Field>
        <Field label="Next Review Date"><TextInput type="date" value={f.nextPasswordChange} onChange={(value) => updateForm("credential", "nextPasswordChange", value)} /></Field>
        {selectedAsset && <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 md:col-span-2 xl:col-span-3"><p className="text-sm font-bold text-blue-900">Linked asset details</p><div className="mt-3 grid gap-3 text-xs text-blue-900 sm:grid-cols-2 xl:grid-cols-5"><p><b>Device:</b><br />{selectedAsset.deviceName || "-"}</p><p><b>Model:</b><br />{selectedAsset.model || "-"}</p><p><b>Serial:</b><br />{selectedAsset.serialNumber || "-"}</p><p><b>IP:</b><br />{selectedAsset.ipAddress || "-"}</p><p><b>Quantity:</b><br />{selectedAsset.quantity || 1}</p></div></div>}
      </>}

      {f.type === "access" && <>
        <Field label="Linked Shared Account" required><SelectInput value={f.linkedCredentialId} onChange={(value) => updateForm("credential", "linkedCredentialId", value)}><option value="">Select shared account</option>{sharedAccounts.map((account) => <option key={account.id} value={account.id}>{account.accountName || account.platform} - {account.scopeType || "Other"}</option>)}</SelectInput></Field>
        <Field label="User Full Name" required><TextInput value={f.owner} onChange={(value) => updateForm("credential", "owner", value)} /></Field>
        <Field label="Job Role"><TextInput value={f.jobRole} onChange={(value) => updateForm("credential", "jobRole", value)} placeholder="Supervisor / Manager / IT Engineer" /></Field>
        <Field label="Department / Operation"><TextInput value={f.departmentOperation} onChange={(value) => updateForm("credential", "departmentOperation", value)} placeholder="QSR / DTF / Stockroom" /></Field>
        <Field label="Email"><TextInput type="email" value={f.email} onChange={(value) => updateForm("credential", "email", value)} /></Field>
        <Field label="Mobile"><TextInput value={f.mobile} onChange={(value) => updateForm("credential", "mobile", value)} /></Field>
        <Field label="Permission Level"><SelectInput value={f.permissionLevel} onChange={(value) => updateForm("credential", "permissionLevel", value)}>{["Viewer", "Operator", "Administrator", "Temporary Access", "Custom"].map((item) => <option key={item}>{item}</option>)}</SelectInput></Field>
        <Field label="Granted Date"><TextInput type="date" value={f.grantedDate} onChange={(value) => updateForm("credential", "grantedDate", value)} /></Field>
        <Field label="Approved By"><TextInput value={f.approvedBy} onChange={(value) => updateForm("credential", "approvedBy", value)} /></Field>
        <Field label="Revoked Date"><TextInput type="date" value={f.revokedDate} onChange={(value) => updateForm("credential", "revokedDate", value)} /></Field>
        {selectedShared && <div className="rounded-2xl border border-violet-200 bg-violet-50 p-4 md:col-span-2 xl:col-span-3"><p className="text-sm font-bold text-violet-900">Assigned shared account</p><p className="mt-1 text-sm text-violet-800">{selectedShared.platform} • {selectedShared.accountName} • {selectedShared.scopeType}{selectedShared.scopeName ? ` / ${selectedShared.scopeName}` : ""}</p></div>}
      </>}

      <label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Notes</span><TextArea value={f.notes} onChange={(value) => updateForm("credential", "notes", value)} placeholder="Purpose, access scope, recovery or review notes" /></label>
      <ModalButtons save={editId ? "Save Vault Changes" : f.type === "access" ? "Save Access Assignment" : "Encrypt & Save Record"} close={close} />
    </form>
  </Modal>;
}

function GroupBrandingModal({ forms, updateForm, submit, formError, close }) {
  const f = forms.groupBranding;
  function uploadLogo(file) {
    if (!file) return;
    if (!file.type || !file.type.startsWith("image/")) return alert("Please upload a PNG/JPG image file.");
    if (file.size > 2 * 1024 * 1024) return alert("Logo file is too large. Please upload below 2 MB.");
    const reader = new FileReader();
    reader.onload = () => updateForm("groupBranding", "logoDataUrl", String(reader.result || ""));
    reader.readAsDataURL(file);
  }
  return <Modal title="Edit Group Branding" subtitle="Logo and footer used automatically for Store Handover forms under this group." onClose={close} max="max-w-3xl">
    {formError && <ErrorBox text={formError} />}
    <form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2">
      <Field label="Group Name" required><TextInput value={f.name} onChange={(v) => updateForm("groupBranding", "name", v)} /></Field>
      <Field label="Footer Color"><TextInput type="color" value={f.footerColor || "#991b1e"} onChange={(v) => updateForm("groupBranding", "footerColor", v)} /></Field>
      <Field label="Footer Phone"><TextInput value={f.footerPhone} onChange={(v) => updateForm("groupBranding", "footerPhone", v)} placeholder="Not configured" /></Field>
      <Field label="Footer Website"><TextInput value={f.footerWebsite} onChange={(v) => updateForm("groupBranding", "footerWebsite", v)} placeholder="www.example.com" /></Field>
      <label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Footer Address</span><TextArea value={f.footerAddress} onChange={(v) => updateForm("groupBranding", "footerAddress", v)} placeholder="Office / company address" /></label>
      <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-4 md:col-span-2">
        <p className="text-sm font-semibold text-slate-700">Group Logo</p>
        <p className="mt-1 text-xs text-slate-500">Used in Store Handover for all branches under this group. Branch logo can still override this.</p>
        <div className="mt-3 flex min-h-24 items-center justify-center rounded-xl bg-white p-3">
          {f.logoDataUrl ? <img src={f.logoDataUrl} alt="Group logo" className="max-h-24 max-w-full object-contain" /> : <span className="text-sm font-bold text-slate-400">No group logo uploaded</span>}
        </div>
        <div className="mt-3 flex flex-wrap gap-2">
          <label className="cursor-pointer rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Upload Logo<input type="file" accept="image/*" onChange={(e) => uploadLogo(e.target.files?.[0])} className="hidden" /></label>
          <button type="button" onClick={() => updateForm("groupBranding", "logoDataUrl", "")} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Remove Logo</button>
        </div>
      </div>
      <div className="rounded-2xl p-4 text-white md:col-span-2" style={{ backgroundColor: f.footerColor || "#991b1e" }}>
        <div className="grid gap-3 text-xs font-semibold md:grid-cols-3">
          <div>☎ {f.footerPhone || "Not configured"}</div>
          <div className="text-center">📍 {f.footerAddress || "Office address"}</div>
          <div className="text-right">◎ {f.footerWebsite || "www.example.com"}</div>
        </div>
      </div>
      <ModalButtons save="Save Group Branding" close={close} />
    </form>
  </Modal>;
}



function TemplateBuilderModal({ templateBuilder, actions, close }) {
  const def = templateDefinitions[templateBuilder.type] || templateDefinitions.asset;
  const selectedCount = def.fields.filter((field) => templateBuilder.selected?.[field.field]).length;
  return <Modal title={def.title} subtitle="Choose required fields, download Excel, fill the sample row, then import it back from the same page." onClose={close} max="max-w-5xl">
    <div className="mt-6 space-y-5">
      <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-800">
        Important fields are selected by default. The downloaded template includes one sample row. Replace the sample values with your actual data before importing.
      </div>
      <div className="flex flex-col gap-3 rounded-2xl border border-slate-200 bg-white p-4 md:flex-row md:items-center md:justify-between">
        <div>
          <p className="text-sm font-bold text-slate-700">Selected Fields</p>
          <p className="text-xs text-slate-500">{selectedCount} of {def.fields.length} field(s) selected.</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button type="button" onClick={() => actions.selectTemplatePreset("important")} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Important Fields</button>
          <button type="button" onClick={() => actions.selectTemplatePreset("all")} className="rounded-2xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-semibold text-indigo-700">Select All</button>
          <button type="button" onClick={() => actions.selectTemplatePreset("none")} className="rounded-2xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700">Clear</button>
        </div>
      </div>
      <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
        {def.fields.map((field) => <label key={field.field} className={`flex items-start gap-3 rounded-2xl border p-3 ${templateBuilder.selected?.[field.field] ? "border-slate-900 bg-slate-50" : "border-slate-200 bg-white"}`}>
          <input type="checkbox" checked={!!templateBuilder.selected?.[field.field]} onChange={(event) => actions.updateTemplateField(field.field, event.target.checked)} className="mt-1 h-4 w-4 accent-slate-900" />
          <span>
            <span className="block text-sm font-bold text-slate-800">{field.label}</span>
            <span className="text-xs text-slate-500">{field.important ? "Important default field" : "Optional field"}</span>
          </span>
        </label>)}
      </div>
      <div className="flex flex-wrap gap-2">
        <button type="button" onClick={actions.downloadSelectedTemplate} className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white">Download Excel Template</button>
        <button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Close</button>
      </div>
    </div>
  </Modal>;
}

function ExcelLocationImportModal({ locationImport, actions, close }) {
  const fileInputRef = useRef(null);
  const hasRows = locationImport.rows.length > 0;
  const headers = locationImport.headers || [];
  const preview = locationImport.preview;
  return <Modal title="Import Locations from Excel" subtitle="Bulk create mall/site locations before importing branches." onClose={close} max="max-w-6xl">
    <div className="mt-6 space-y-5">
      <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-800">
        Upload your location template, select the sheet, map required columns, preview, then confirm import. Duplicate check uses Location / Mall + City + Emirate.
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <div className="rounded-2xl border border-slate-200 bg-white p-4 md:col-span-2">
          <p className="text-sm font-bold text-slate-700">Excel File</p>
          <p className="mt-1 text-xs text-slate-500">Supported: .xlsx, .xls, .csv</p>
          <div className="mt-3 flex flex-wrap gap-2">
            <button type="button" onClick={() => fileInputRef.current?.click()} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Choose Location Excel File</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => actions.handleLocationExcelFile(event.target.files?.[0])} />
            {locationImport.fileName && <span className="rounded-2xl bg-slate-100 px-4 py-2 text-sm font-semibold text-slate-700">{locationImport.fileName}</span>}
          </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-4">
          <p className="text-sm font-bold text-slate-700">Duplicate Locations</p>
          <select value={locationImport.duplicateMode} onChange={(event) => actions.updateLocationImportDuplicateMode(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
            <option value="skip">Skip duplicates</option>
            <option value="update">Update existing locations</option>
            <option value="import">Import anyway</option>
          </select>
        </div>
      </div>

      {locationImport.sheetNames.length > 0 && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <p className="text-sm font-bold text-slate-700">Select Sheet</p>
        <select value={locationImport.selectedSheet} onChange={(event) => actions.loadLocationExcelSheet(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
          {locationImport.sheetNames.map((sheet) => <option key={sheet}>{sheet}</option>)}
        </select>
      </div>}

      {hasRows && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
          <div><p className="text-sm font-bold text-slate-700">Column Mapping</p><p className="text-xs text-slate-500">Minimum required: Location / Mall, City and Emirate.</p></div>
          <button type="button" onClick={() => actions.buildLocationImportPreview(true)} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Preview Import</button>
        </div>
        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {locationImportFields.map(([field, label, required]) => <Field key={field} label={`${label}${required ? " *" : ""}`}>
            <select value={locationImport.mapping[field] || ""} onChange={(event) => actions.updateLocationImportMapping(field, event.target.value)} className="w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
              <option value="">Do not import</option>
              {headers.map((header) => <option key={header} value={header}>{header}</option>)}
            </select>
          </Field>)}
        </div>
      </div>}

      {preview && <div className="space-y-4 rounded-2xl border border-slate-200 bg-white p-4">
        <div className="grid gap-3 md:grid-cols-6">
          <Card title="Rows" value={preview.totalRows} note="Excel rows" />
          <Card title="New" value={preview.newCount} note="New locations" />
          <Card title="Duplicates" value={preview.duplicates} note="Matched existing" />
          <Card title="Skipped" value={preview.skipCount} note="Duplicate skip" />
          <Card title="Missing" value={preview.missingRequired} note="Required missing" />
          <Card title="To Save" value={preview.importable.length} note="After rules" />
        </div>
        <div className="max-h-72 overflow-auto rounded-2xl border border-slate-200">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-600"><tr><th className="px-3 py-2">Row</th><th className="px-3 py-2">Location</th><th className="px-3 py-2">City</th><th className="px-3 py-2">Emirate</th><th className="px-3 py-2">Type</th><th className="px-3 py-2">Duplicate</th></tr></thead>
            <tbody className="divide-y divide-slate-200">{preview.parsedLocations.slice(0, 80).map((location) => <tr key={`${location.sourceRow}-${location.name}`}><td className="px-3 py-2">{location.sourceRow}</td><td className="px-3 py-2 font-semibold">{location.name}</td><td className="px-3 py-2">{location.city}</td><td className="px-3 py-2">{location.emirate}</td><td className="px-3 py-2">{location.type}</td><td className="px-3 py-2">{location.duplicate ? "Yes" : "No"}</td></tr>)}</tbody>
          </table>
        </div>
        <div className="flex flex-wrap gap-2"><button type="button" onClick={actions.confirmLocationImport} className="rounded-2xl bg-green-700 px-5 py-2.5 text-sm font-semibold text-white">Confirm Import Locations</button><button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button></div>
      </div>}
    </div>
  </Modal>;
}

function ExcelBranchImportModal({ branchImport, actions, close }) {
  const fileInputRef = useRef(null);
  const hasRows = branchImport.rows.length > 0;
  const headers = branchImport.headers || [];
  const preview = branchImport.preview;
  return <Modal title="Import Branches from Excel" subtitle="Bulk create 70+ branch records from your Excel template. Remaining details can be updated later." onClose={close} max="max-w-6xl">
    <div className="mt-6 space-y-5">
      <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-800">
        Upload your branch template, select the sheet, map only required columns, preview, then confirm import. Missing locations and new concepts will be created automatically.
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <div className="rounded-2xl border border-slate-200 bg-white p-4 md:col-span-2">
          <p className="text-sm font-bold text-slate-700">Excel File</p>
          <p className="mt-1 text-xs text-slate-500">Supported: .xlsx, .xls, .csv</p>
          <div className="mt-3 flex flex-wrap gap-2">
            <button type="button" onClick={() => fileInputRef.current?.click()} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Choose Branch Excel File</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => actions.handleBranchExcelFile(event.target.files?.[0])} />
            {branchImport.fileName && <span className="rounded-2xl bg-slate-100 px-4 py-2 text-sm font-semibold text-slate-700">{branchImport.fileName}</span>}
          </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-4">
          <p className="text-sm font-bold text-slate-700">Duplicate Branches</p>
          <select value={branchImport.duplicateMode} onChange={(event) => actions.updateBranchImportDuplicateMode(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
            <option value="skip">Skip duplicates</option>
            <option value="update">Update existing branches</option>
            <option value="import">Import anyway</option>
          </select>
          <p className="mt-2 text-xs text-slate-500">Duplicate check uses Branch Display Name + Location / Mall.</p>
        </div>
      </div>

      {branchImport.sheetNames.length > 0 && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <p className="text-sm font-bold text-slate-700">Select Sheet</p>
        <select value={branchImport.selectedSheet} onChange={(event) => actions.loadBranchExcelSheet(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
          {branchImport.sheetNames.map((sheet) => <option key={sheet}>{sheet}</option>)}
        </select>
      </div>}

      {hasRows && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
          <div><p className="text-sm font-bold text-slate-700">Column Mapping</p><p className="text-xs text-slate-500">Map only the fields you need. Required fields are marked with *.</p></div>
          <button type="button" onClick={() => actions.buildBranchImportPreview(true)} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Preview Import</button>
        </div>
        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {branchImportFields.map(([field, label, required]) => <Field key={field} label={`${label}${required ? " *" : ""}`}>
            <select value={branchImport.mapping[field] || ""} onChange={(event) => actions.updateBranchImportMapping(field, event.target.value)} className="w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
              <option value="">Do not import</option>
              {headers.map((header) => <option key={header} value={header}>{header}</option>)}
            </select>
          </Field>)}
        </div>
      </div>}

      {preview && <div className="space-y-4 rounded-2xl border border-slate-200 bg-white p-4">
        <div className="grid gap-3 md:grid-cols-6">
          <Card title="Rows" value={preview.totalRows} note="Excel rows" />
          <Card title="New" value={preview.newCount} note="New branches" />
          <Card title="Duplicates" value={preview.duplicates} note="Matched existing" />
          <Card title="Skipped" value={preview.skipCount} note="Duplicate skip" />
          <Card title="Missing" value={preview.missingRequired} note="Required missing" />
          <Card title="To Save" value={preview.importable.length} note="After rules" />
        </div>
        <div className="max-h-72 overflow-auto rounded-2xl border border-slate-200">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-600"><tr><th className="px-3 py-2">Row</th><th className="px-3 py-2">Group</th><th className="px-3 py-2">Concept</th><th className="px-3 py-2">Branch</th><th className="px-3 py-2">Location</th><th className="px-3 py-2">Type</th><th className="px-3 py-2">Status</th><th className="px-3 py-2">Duplicate</th></tr></thead>
            <tbody className="divide-y divide-slate-200">{preview.parsedBranches.slice(0, 60).map((branch) => <tr key={`${branch.sourceRow}-${branch.branchName}`}><td className="px-3 py-2">{branch.sourceRow}</td><td className="px-3 py-2">{branch.groupName}</td><td className="px-3 py-2">{branch.conceptName}</td><td className="px-3 py-2 font-semibold">{branch.branchName}</td><td className="px-3 py-2">{branch.locationName}</td><td className="px-3 py-2">{branch.branchType}</td><td className="px-3 py-2">{branch.status}</td><td className="px-3 py-2">{branch.duplicate ? "Yes" : "No"}</td></tr>)}</tbody>
          </table>
        </div>
        <div className="flex flex-wrap gap-2"><button type="button" onClick={actions.confirmBranchImport} className="rounded-2xl bg-green-700 px-5 py-2.5 text-sm font-semibold text-white">Confirm Import Branches</button><button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button></div>
      </div>}
    </div>
  </Modal>;
}

function ExcelProductImportModal({ productImport, actions, close }) {
  const fileInputRef = useRef(null);
  const hasRows = productImport.rows.length > 0;
  const headers = productImport.headers || [];
  const preview = productImport.preview;
  return <Modal title="Upload Asset Product List" subtitle="Import multiple reusable products using the downloaded Product List Excel template." onClose={close} max="max-w-6xl">
    <div className="mt-6 space-y-5">
      <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-800">
        Recommended process: download the Product Template, complete the rows, upload the same file here, preview the records, then confirm import.
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <div className="rounded-2xl border border-slate-200 bg-white p-4 md:col-span-2">
          <p className="text-sm font-bold text-slate-700">Product Excel File</p>
          <p className="mt-1 text-xs text-slate-500">Supported formats: .xlsx, .xls and .csv</p>
          <div className="mt-3 flex flex-wrap gap-2">
            <button type="button" onClick={() => fileInputRef.current?.click()} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Choose Product Excel</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => actions.handleProductExcelFile(event.target.files?.[0])} />
            {productImport.fileName && <span className="rounded-2xl bg-slate-100 px-4 py-2 text-sm font-semibold text-slate-700">{productImport.fileName}</span>}
          </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-4">
          <p className="text-sm font-bold text-slate-700">Duplicate Products</p>
          <select value={productImport.duplicateMode} onChange={(event) => actions.updateProductImportDuplicateMode(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
            <option value="skip">Skip existing products</option>
            <option value="update">Update existing products</option>
            <option value="import">Import another copy</option>
          </select>
          <p className="mt-2 text-xs text-slate-500">Products are matched mainly by Product Name.</p>
        </div>
      </div>

      {productImport.sheetNames.length > 0 && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <p className="text-sm font-bold text-slate-700">Select Excel Sheet</p>
        <select value={productImport.selectedSheet} onChange={(event) => actions.loadProductExcelSheet(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
          {productImport.sheetNames.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
        </select>
        <p className="mt-2 text-xs text-slate-500">Rows found: {productImport.rows.length}</p>
      </div>}

      {hasRows && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
          <div><p className="text-sm font-bold text-slate-700">Column Mapping</p><p className="text-xs text-slate-500">The downloaded template should map automatically. Adjust only when your Excel uses different column names.</p></div>
          <button type="button" onClick={() => actions.buildProductImportPreview(true)} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Preview Products</button>
        </div>
        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {productMasterFields.map(([field, label, required]) => <label key={field} className="block">
            <span className="text-xs font-bold text-slate-600">{label} {required && <span className="text-red-500">*</span>}</span>
            <select value={productImport.mapping[field] || ""} onChange={(event) => actions.updateProductImportMapping(field, event.target.value)} className="mt-1 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
              <option value="">{required ? "Select column" : "Do not import"}</option>
              {headers.map((header) => <option key={`${field}-${header}`} value={header}>{header}</option>)}
            </select>
          </label>)}
        </div>
      </div>}

      {preview && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="grid gap-3 md:grid-cols-5">
          <InfoPill label="Excel Rows" value={preview.totalRows} />
          <InfoPill label="Valid Products" value={preview.parsedProducts.length} />
          <InfoPill label="To Save" value={preview.importable.length} />
          <InfoPill label="Duplicates" value={preview.duplicates} />
          <InfoPill label="Invalid Rows" value={preview.invalid} />
        </div>
        <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200">
          <table className="min-w-[900px] w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-600"><tr><th className="px-3 py-2">Row</th><th className="px-3 py-2">Product</th><th className="px-3 py-2">Category</th><th className="px-3 py-2">Brand</th><th className="px-3 py-2">Model</th><th className="px-3 py-2">Qty</th><th className="px-3 py-2">Default Location</th><th className="px-3 py-2">Result</th></tr></thead>
            <tbody className="divide-y divide-slate-200">{preview.parsedProducts.slice(0, 15).map((product) => <tr key={`${product.sourceRow}-${product.productName}`}><td className="px-3 py-2">{product.sourceRow}</td><td className="px-3 py-2 font-semibold">{product.productName}</td><td className="px-3 py-2">{product.category}</td><td className="px-3 py-2">{product.brand || "-"}</td><td className="px-3 py-2">{product.model || "-"}</td><td className="px-3 py-2">{product.defaultQuantity}</td><td className="px-3 py-2">{product.defaultLocation || "-"}</td><td className="px-3 py-2">{product.duplicate ? <Badge type="warning">Existing</Badge> : <Badge type="success">New</Badge>}</td></tr>)}</tbody>
          </table>
        </div>
        {preview.parsedProducts.length > 15 && <p className="mt-2 text-xs text-slate-500">Showing the first 15 valid products. All valid rows will be processed.</p>}
      </div>}

      <div className="flex flex-wrap gap-2">
        <button type="button" onClick={actions.confirmProductImport} className="rounded-2xl bg-green-700 px-5 py-2.5 text-sm font-semibold text-white">Confirm Product Import</button>
        <button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button>
      </div>
    </div>
  </Modal>;
}

function ExcelAssetImportModal({ excelImport, actions, selectedBranch, close }) {
  const fileInputRef = useRef(null);
  const hasRows = excelImport.rows.length > 0;
  const headers = excelImport.headers || [];
  const preview = excelImport.preview;
  return <Modal title="Import Assets from Excel" subtitle={`Selected branch: ${selectedBranch.branchName} - ${selectedBranch.locationName}, ${selectedBranch.city}`} onClose={close} max="max-w-6xl">
    <div className="mt-6 space-y-5">
      <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-800">
        Upload your Excel file, select the correct sheet, map only the columns you want, preview, then confirm import. Location and Department columns are ignored unless you map them manually.
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <div className="rounded-2xl border border-slate-200 bg-white p-4 md:col-span-2">
          <p className="text-sm font-bold text-slate-700">Excel File</p>
          <p className="mt-1 text-xs text-slate-500">Supported: .xlsx, .xls, .csv</p>
          <div className="mt-3 flex flex-wrap gap-2">
            <button type="button" onClick={() => fileInputRef.current?.click()} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Choose Excel File</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => actions.handleExcelAssetFile(event.target.files?.[0])} />
            {excelImport.fileName && <span className="rounded-2xl bg-slate-100 px-4 py-2 text-sm font-semibold text-slate-700">{excelImport.fileName}</span>}
          </div>
        </div>
        <div className="rounded-2xl border border-slate-200 bg-white p-4">
          <p className="text-sm font-bold text-slate-700">Duplicate Serial Numbers</p>
          <select value={excelImport.duplicateMode} onChange={(event) => actions.updateExcelDuplicateMode(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
            <option value="skip">Skip duplicates</option>
            <option value="update">Update existing assets</option>
            <option value="import">Import anyway</option>
          </select>
        </div>
      </div>

      {excelImport.sheetNames.length > 0 && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <p className="text-sm font-bold text-slate-700">Select Excel Sheet</p>
        <select value={excelImport.selectedSheet} onChange={(event) => actions.loadExcelSheet(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
          {excelImport.sheetNames.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
        </select>
        <p className="mt-2 text-xs text-slate-500">Rows found in selected sheet: {excelImport.rows.length}</p>
      </div>}

      {hasRows && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
          <div>
            <p className="text-sm font-bold text-slate-700">Column Mapping</p>
            <p className="text-xs text-slate-500">Map only the required/useful columns. Unwanted columns like Location and Department can remain unused.</p>
          </div>
          <button type="button" onClick={() => actions.buildExcelImportPreview(true)} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">Preview Import</button>
        </div>
        <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-4">
          {excelImportFields.map(([field, label, required]) => <label key={field} className="block">
            <span className="text-xs font-bold text-slate-600">{label} {required && <span className="text-red-500">*</span>}</span>
            <select value={excelImport.mapping[field] || ""} onChange={(event) => actions.updateExcelMapping(field, event.target.value)} className="mt-1 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none">
              <option value="">{required ? "Select column" : "Do not import"}</option>
              {headers.map((header) => <option key={`${field}-${header}`} value={header}>{header}</option>)}
            </select>
          </label>)}
        </div>
      </div>}

      {preview && <div className="rounded-2xl border border-slate-200 bg-white p-4">
        <div className="grid gap-3 md:grid-cols-5">
          <InfoPill label="Rows" value={preview.totalRows} />
          <InfoPill label="Valid Assets" value={preview.parsedAssets.length} />
          <InfoPill label="To Save" value={preview.importable.length} />
          <InfoPill label="Duplicates" value={preview.duplicates} />
          <InfoPill label="No Serial" value={preview.missingSerial} />
        </div>
        <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200">
          <table className="min-w-[900px] w-full text-left text-xs">
            <thead className="bg-slate-100 text-slate-600"><tr><th className="px-3 py-2">Row</th><th className="px-3 py-2">Category</th><th className="px-3 py-2">Asset Name</th><th className="px-3 py-2">Serial</th><th className="px-3 py-2">Description</th><th className="px-3 py-2">Brand</th><th className="px-3 py-2">Model</th><th className="px-3 py-2">Qty</th><th className="px-3 py-2">Status</th></tr></thead>
            <tbody className="divide-y divide-slate-200">{preview.parsedAssets.slice(0, 12).map((asset) => <tr key={`${asset.sourceRow}-${asset.serialNumber}`}><td className="px-3 py-2">{asset.sourceRow}</td><td className="px-3 py-2 font-semibold">{asset.category}</td><td className="px-3 py-2">{asset.deviceName}</td><td className="px-3 py-2">{asset.serialNumber}</td><td className="px-3 py-2">{asset.description}</td><td className="px-3 py-2">{asset.brand}</td><td className="px-3 py-2">{asset.model}</td><td className="px-3 py-2">{asset.quantity}</td><td className="px-3 py-2">{asset.duplicate ? <Badge type="warning">Duplicate</Badge> : <Badge type="success">New</Badge>}</td></tr>)}</tbody>
          </table>
        </div>
        {preview.parsedAssets.length > 12 && <p className="mt-2 text-xs text-slate-500">Showing first 12 rows only. All valid rows will be processed based on duplicate mode.</p>}
      </div>}

      <div className="flex flex-wrap gap-2">
        <button type="button" onClick={actions.confirmExcelAssetImport} className="rounded-2xl bg-green-700 px-5 py-2.5 text-sm font-semibold text-white">Confirm Import to Branch</button>
        <button type="button" onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button>
      </div>
    </div>
  </Modal>;
}

function InfoPill({ label, value }) {
  return <div className="rounded-2xl bg-slate-50 p-3 text-center"><p className="text-xs font-semibold uppercase text-slate-500">{label}</p><p className="mt-1 text-2xl font-bold text-slate-900">{value}</p></div>;
}

function AssetModal({ forms, updateForm, submit, formError, close, editId, assetProducts = [], assetCategories = [] }) {
  const f = forms.asset;
  const categories = assetCategories.filter((item) => item.isActive !== false).map((item) => item.name).sort();
  const applyProduct = (productId) => {
    const product = assetProducts.find((item) => String(item.id) === String(productId));
    if (!product) return;
    updateForm("asset", "category", product.category || "Other IT Device");
    updateForm("asset", "deviceName", product.productName || "");
    updateForm("asset", "description", product.description || "");
    updateForm("asset", "quantity", product.defaultQuantity || "1");
    updateForm("asset", "brand", product.brand || "");
    updateForm("asset", "model", product.model || "");
    updateForm("asset", "installedLocation", product.defaultLocation || "");
    updateForm("asset", "status", product.defaultStatus || "Installed");
    if (product.remarks) updateForm("asset", "remarks", cleanImportedPlaceholder(product.remarks));
  };
  const activeProducts = assetProducts.filter((item) => item.isActive !== false);
  const productCategories = Array.from(new Set(activeProducts.map((item) => item.category).filter(Boolean))).sort();
  const filteredProducts = activeProducts.filter((item) => item.category === f.category);
  return <Modal title={editId ? "Edit IT Asset" : "Add IT Asset"} subtitle="Select asset type and model/product to auto-fill common fields. Then update serial, IP, MAC and warranty." onClose={close} max="max-w-6xl">{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
    {activeProducts.length > 0 && <div className="rounded-2xl border border-indigo-200 bg-indigo-50 p-4 md:col-span-2 xl:col-span-3"><div><p className="text-sm font-bold text-indigo-900">Smart Asset Product Auto-Fill</p><p className="mt-1 text-xs text-indigo-700">Choose asset type first. The model/product dropdown will show matching saved products only.</p></div><div className="mt-3 grid gap-3 md:grid-cols-2"><Field label="Asset Type"><SelectInput value={f.category} onChange={(v) => updateForm("asset", "category", v)}>{[...new Set([...categories, ...productCategories])].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Saved Product / Model"><select onChange={(event) => applyProduct(event.target.value)} defaultValue="" className="w-full rounded-xl border border-indigo-200 bg-white px-4 py-2.5 text-sm font-semibold text-indigo-800 outline-none"><option value="">Select matching product/model</option>{(filteredProducts.length ? filteredProducts : activeProducts).map((product) => <option key={product.id} value={product.id}>{product.productName} {product.model ? `- ${product.model}` : ""}</option>)}</select></Field></div></div>}
    <Field label="Device Category" required><SelectInput value={f.category} onChange={(v) => updateForm("asset", "category", v)}>{categories.map((x) => <option key={x}>{x}</option>)}</SelectInput></Field>
    {f.category === "Other IT Device" && <Field label="Custom Category Name" required><TextInput value={f.customCategory} onChange={(v) => updateForm("asset", "customCategory", v)} /></Field>}
    <Field label="Device Name / Asset Name" required><TextInput value={f.deviceName} onChange={(v) => updateForm("asset", "deviceName", v)} /></Field>
    <Field label="Serial Number" required><TextInput value={f.serialNumber} onChange={(v) => updateForm("asset", "serialNumber", v)} /></Field>
    <Field label="Brand"><TextInput value={f.brand} onChange={(v) => updateForm("asset", "brand", v)} /></Field>
    <Field label="Model"><TextInput value={f.model} onChange={(v) => updateForm("asset", "model", v)} /></Field>
    <Field label="Quantity"><TextInput type="number" value={f.quantity || "1"} onChange={(v) => updateForm("asset", "quantity", v)} /></Field>
    <Field label="IP Address"><TextInput value={f.ipAddress} onChange={(v) => updateForm("asset", "ipAddress", v)} /></Field>
    <Field label="MAC Address"><TextInput value={f.macAddress} onChange={(v) => updateForm("asset", "macAddress", v)} /></Field>
    <Field label="Installed Location" required><TextInput value={f.installedLocation} onChange={(v) => updateForm("asset", "installedLocation", v)} /></Field>
    <Field label="Switch Port"><TextInput value={f.switchPort} onChange={(v) => updateForm("asset", "switchPort", v)} /></Field>
    {f.category === "CCTV Camera" && <><Field label="Camera Channel"><TextInput value={f.channelNumber} onChange={(v) => updateForm("asset", "channelNumber", v)} /></Field><Field label="Camera View / Area"><TextInput value={f.cameraView} onChange={(v) => updateForm("asset", "cameraView", v)} /></Field></>}
    <Field label="Warranty Start"><TextInput type="date" value={f.warrantyStart} onChange={(v) => updateForm("asset", "warrantyStart", v)} /></Field>
    <Field label="Warranty End"><TextInput type="date" value={f.warrantyEnd} onChange={(v) => updateForm("asset", "warrantyEnd", v)} /></Field>
    <Field label="Status" required><SelectInput value={f.status} onChange={(v) => updateForm("asset", "status", v)}>{["Installed", "Tested", "Online", "Pending Configuration", "Faulty", "Replaced", "Spare"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field>
    <label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Description <span className="font-normal text-slate-400">(optional)</span></span><TextArea value={f.description} onChange={(v) => updateForm("asset", "description", v)} placeholder="Example: SENOR F5/K5H 15 inch touch screen POS system, IP54 Rating" /></label>
    <label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Remarks</span><TextArea value={f.remarks} onChange={(v) => updateForm("asset", "remarks", v)} /></label>
    <ModalButtons save={editId ? "Update IT Asset" : "Save IT Asset"} close={close} />
  </form></Modal>;
}

function AssetTemplateModal({ forms, updateForm, submit, formError, close, editId, assetProducts = [] }) {
  const f = forms.assetTemplate;
  const rows = Array.isArray(f.rows) && f.rows.length ? f.rows : [{ ...emptyTemplateRow }];
  const updateRow = (index, field, value) => {
    const next = rows.map((row, i) => i === index ? { ...row, [field]: value } : row);
    if (field === "productId") {
      const product = assetProducts.find((item) => String(item.id) === String(value));
      if (product) {
        next[index].quantity = next[index].quantity || product.defaultQuantity || "1";
        next[index].installedLocation = next[index].installedLocation || product.defaultLocation || "";
      }
    }
    updateForm("assetTemplate", "rows", next);
  };
  const addRow = () => updateForm("assetTemplate", "rows", [...rows, { ...emptyTemplateRow }]);
  const removeRow = (index) => updateForm("assetTemplate", "rows", rows.filter((_, i) => i !== index));
  return <Modal title={editId ? "Edit Branch Asset Template" : "Create Branch Asset Template"} subtitle="Build your own standard setup. Add POS, router, switch, NVR, cameras, printers and other products with default quantity/location." onClose={close} max="max-w-6xl">{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 space-y-5">
    <div className="grid gap-4 md:grid-cols-3"><Field label="Template Name" required><TextInput value={f.name} onChange={(v) => updateForm("assetTemplate", "name", v)} placeholder="Example: Standard QSR Branch Setup" /></Field><Field label="Branch Type"><SelectInput value={f.branchType} onChange={(v) => updateForm("assetTemplate", "branchType", v)}>{["QSR", "Dining", "DTF", "Cafe", "Accommodation", "Warehouse", "Central Kitchen", "Office", "Other"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Description"><TextInput value={f.description} onChange={(v) => updateForm("assetTemplate", "description", v)} placeholder="Short note" /></Field></div>
    <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><div className="flex items-center justify-between gap-3"><div><p className="font-bold">Template Product Rows</p><p className="text-xs text-slate-500">Each row becomes one asset line in the branch. Quantity can be edited later.</p></div><button type="button" onClick={addRow} className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">+ Add Row</button></div><div className="mt-4 space-y-3">{rows.map((row, index) => { const product = assetProducts.find((item) => String(item.id) === String(row.productId)); return <div key={index} className="grid gap-3 rounded-2xl border border-slate-200 bg-white p-3 md:grid-cols-[2fr_90px_1fr_1fr_auto]"><Field label="Product"><SelectInput value={row.productId} onChange={(v) => updateRow(index, "productId", v)}><option value="">Select product</option>{assetProducts.filter((product) => product.isActive !== false).map((product) => <option key={product.id} value={product.id}>{product.category} - {product.productName} {product.model ? `(${product.model})` : ""}</option>)}</SelectInput>{product && <p className="mt-1 text-xs text-slate-500">{product.brand || "Brand blank"} • {product.model || "Model blank"}</p>}</Field><Field label="Qty"><TextInput type="number" value={row.quantity || "1"} onChange={(v) => updateRow(index, "quantity", v)} /></Field><Field label="Installed Location"><TextInput value={row.installedLocation} onChange={(v) => updateRow(index, "installedLocation", v)} placeholder={product?.defaultLocation || "IT Rack"} /></Field><Field label="Remarks"><TextInput value={row.remarks} onChange={(v) => updateRow(index, "remarks", v)} /></Field><div className="flex items-end"><button type="button" onClick={() => removeRow(index)} disabled={rows.length === 1} className="rounded-xl border border-red-200 px-3 py-2 text-xs font-semibold text-red-700 disabled:opacity-40">Remove</button></div></div>; })}</div></div>
    <ModalButtons save={editId ? "Update Branch Template" : "Save Branch Template"} close={close} />
  </form></Modal>;
}

function ApplyAssetTemplateModal({ close, selectedBranch, assetTemplates = [], assetProducts = [], applyBranchAssetTemplate }) {
  const [templateId, setTemplateId] = useState(String(assetTemplates[0]?.id || ""));
  const template = assetTemplates.find((item) => String(item.id) === String(templateId));
  return <Modal title="Apply Standard Asset Template" subtitle={`Branch: ${selectedBranch?.branchName || "Selected Branch"}. This will add asset rows from the selected template.`} onClose={close} max="max-w-4xl"><div className="mt-6 space-y-4"><Field label="Select Template" required><SelectInput value={templateId} onChange={setTemplateId}><option value="">Select template</option>{assetTemplates.map((item) => <option key={item.id} value={String(item.id)}>{item.name} - {item.branchType}</option>)}</SelectInput></Field>{template && <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="font-bold">Preview Items</p><p className="mt-1 text-sm text-slate-500">{template.description}</p><div className="mt-3 space-y-2">{(template.rows || []).map((row, index) => { const product = assetProducts.find((item) => String(item.id) === String(row.productId)); return <div key={index} className="rounded-xl bg-white p-3 text-sm"><span className="font-semibold">{product?.productName || "Missing product"}</span><span className="text-slate-500"> • Qty {row.quantity || 1} • {row.installedLocation || product?.defaultLocation || "Location blank"}</span></div>; })}</div></div>}<div className="flex flex-wrap gap-2"><button onClick={() => applyBranchAssetTemplate(templateId)} disabled={!templateId} className="rounded-2xl bg-slate-900 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">Apply Template to Branch</button><button onClick={close} className="rounded-2xl border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700">Cancel</button></div></div></Modal>;
}

function AssetProductModal({ forms, updateForm, submit, formError, close, editId, assetCategories = [], addAssetCategory }) {
  const f = forms.product;
  const [newCategory, setNewCategory] = useState("");
  const [categoryNote, setCategoryNote] = useState("");
  const categories = assetCategories.filter((item) => item.isActive !== false).map((item) => item.name).sort();
  function saveCategory() {
    const result = addAssetCategory(newCategory, categoryNote);
    if (!result?.ok) return alert(result?.message || "Could not add category.");
    updateForm("product", "category", result.category.name);
    setNewCategory(""); setCategoryNote("");
  }
  return <Modal title={editId ? "Edit Asset Product" : "Add Asset Product"} subtitle="Create reusable products for smart asset entry and branch templates." onClose={close} max="max-w-5xl">{formError && <ErrorBox text={formError}/>}<form onSubmit={submit} className="mt-6 space-y-4">
    <div className="rounded-2xl border border-slate-300 bg-slate-50 p-4"><h4 className="font-bold text-slate-900">Basic Product Information</h4><div className="mt-4 grid gap-4 md:grid-cols-2"><Field label="Product Name" required><TextInput value={f.productName} onChange={(value) => updateForm("product","productName",value)} placeholder="Example: NVR - Hikvision DS-7608NI-Q2"/></Field><Field label="Device Category" required><SelectInput value={f.category} onChange={(value) => updateForm("product","category",value)}>{categories.map((category) => <option key={category}>{category}</option>)}</SelectInput></Field><Field label="Brand"><TextInput value={f.brand} onChange={(value) => updateForm("product","brand",value)}/></Field><Field label="Model"><TextInput value={f.model} onChange={(value) => updateForm("product","model",value)}/></Field></div>{f.category === "Other IT Device" && <div className="mt-4 rounded-2xl border border-amber-200 bg-amber-50 p-4"><p className="text-sm font-bold text-amber-900">Create a new reusable category</p><p className="mt-1 text-xs text-amber-700">Add it once and it will become available in all category dropdowns.</p><div className="mt-3 grid gap-3 md:grid-cols-[1fr_1.5fr_auto]"><TextInput value={newCategory} onChange={setNewCategory} placeholder="New category name"/><TextInput value={categoryNote} onChange={setCategoryNote} placeholder="Description (optional)"/><button type="button" onClick={saveCategory} className="rounded-xl bg-amber-500 px-4 py-2 text-sm font-semibold text-slate-950">Save Category</button></div></div>}</div>
    <div className="rounded-2xl border border-slate-300 bg-slate-50 p-4"><h4 className="font-bold text-slate-900">Default Asset Values</h4><div className="mt-4 grid gap-4 md:grid-cols-2"><Field label="Default Quantity"><TextInput type="number" value={f.defaultQuantity||"1"} onChange={(value) => updateForm("product","defaultQuantity",value)}/></Field><Field label="Default Installed Location"><TextInput value={f.defaultLocation} onChange={(value) => updateForm("product","defaultLocation",value)} placeholder="IT Rack / Cash Counter / Kitchen"/></Field><Field label="Default Status"><SelectInput value={f.defaultStatus} onChange={(value) => updateForm("product","defaultStatus",value)}>{["Installed","Tested","Online","Pending Configuration","Faulty","Replaced","Spare"].map((status) => <option key={status}>{status}</option>)}</SelectInput></Field><Field label="Product Status"><SelectInput value={f.isActive === false ? "Inactive" : "Active"} onChange={(value) => updateForm("product","isActive",value === "Active")}><option>Active</option><option>Inactive</option></SelectInput></Field></div></div>
    <div className="rounded-2xl border border-slate-300 bg-slate-50 p-4"><h4 className="font-bold text-slate-900">Additional Information</h4><div className="mt-4 grid gap-4 md:grid-cols-2"><label className="block"><span className="text-sm font-semibold text-slate-700">Description</span><TextArea value={f.description} onChange={(value) => updateForm("product","description",value)}/></label><label className="block"><span className="text-sm font-semibold text-slate-700">Remarks</span><TextArea value={f.remarks} onChange={(value) => updateForm("product","remarks",value)}/></label></div></div>
    <ModalButtons save={editId ? "Update Product" : "Save Product"} close={close}/>
  </form></Modal>;
}
function NetworkModal({ forms, updateForm, submit, formError, close, selectedBranch }) { const f = forms.network; return <Modal title="Add / Edit Network Setup" subtitle={`Branch: ${selectedBranch.branchName} - ${selectedBranch.locationName}`} onClose={close}>{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3">{[["ISP", "isp", true], ["Internet Account", "internetAccount"], ["Bandwidth", "bandwidth"], ["Router Model", "routerModel"], ["Router / Gateway IP", "routerIp", true], ["Switch Model", "switchModel"], ["NVR IP", "nvrIp"], ["CCTV IP Range", "cctvRange"], ["POS IP Range", "posRange"], ["Biometric IP", "biometricIp"], ["Wi-Fi Name", "wifiName"], ["Wi-Fi Password", "wifiPassword"], ["VLAN Details", "vlanDetails"]].map(([label, key, req]) => <Field key={key} label={label} required={req}><TextInput value={f[key]} onChange={(v) => updateForm("network", key, v)} /></Field>)}<label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Remarks</span><TextArea value={f.remarks} onChange={(v) => updateForm("network", "remarks", v)} /></label><ModalButtons save="Save Network Setup" close={close} /></form></Modal>; }
function FaultModal({ forms, updateForm, submit, formError, close, branches, editId }) { const f = forms.fault; return <Modal title={editId ? "Edit Fault Log" : "Add Fault Log"} subtitle="Record what happened, who attended, and what action was taken." onClose={close}>{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3"><Field label="Branch / Location" required><SelectInput value={f.branchId} onChange={(v) => updateForm("fault", "branchId", v)}>{branches.map((b) => <option key={b.id} value={String(b.id)}>{b.branchName} - {b.locationName}, {b.city}</option>)}</SelectInput></Field><Field label="Fault Date" required><TextInput type="date" value={f.faultDate} onChange={(v) => updateForm("fault", "faultDate", v)} /></Field><Field label="Fault Time" required><TextInput type="time" value={f.faultTime} onChange={(v) => updateForm("fault", "faultTime", v)} /></Field><Field label="Category" required><SelectInput value={f.category} onChange={(v) => updateForm("fault", "category", v)}>{["CCTV", "Network", "POS", "Internet", "Biometric", "Door Alarm", "Other IT"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Fault Title" required><TextInput value={f.title} onChange={(v) => updateForm("fault", "title", v)} /></Field><Field label="Reported By"><TextInput value={f.reportedBy} onChange={(v) => updateForm("fault", "reportedBy", v)} /></Field><Field label="Attended By" required><TextInput value={f.attendedBy} onChange={(v) => updateForm("fault", "attendedBy", v)} /></Field><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Fault Description</span><TextArea value={f.description} onChange={(v) => updateForm("fault", "description", v)} /></label><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Action Taken *</span><TextArea value={f.actionTaken} onChange={(v) => updateForm("fault", "actionTaken", v)} /></label><Field label="Status" required><SelectInput value={f.status} onChange={(v) => updateForm("fault", "status", v)}>{["Open", "In Progress", "Waiting Vendor", "Waiting Parts", "Closed"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><Field label="Closed Date"><TextInput type="date" value={f.closedDate} onChange={(v) => updateForm("fault", "closedDate", v)} /></Field><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Remarks</span><TextArea value={f.remarks} onChange={(v) => updateForm("fault", "remarks", v)} /></label><ModalButtons save={editId ? "Update Fault Log" : "Save Fault Log"} close={close} /></form></Modal>; }
function HandoverModal({ type, forms, updateForm, submit, formError, close, assets, selectedBranch }) { const isPullout = type === "pullout"; const f = isPullout ? forms.pullout : forms.handover; const name = isPullout ? "pullout" : "handover"; return <Modal title={isPullout ? "IT Pullout Form" : "IT Handover Form"} subtitle={`Branch: ${selectedBranch.branchName} - ${selectedBranch.locationName}. Asset list will be included automatically.`} onClose={close}>{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3">{isPullout ? <><Field label="Pullout Date" required><TextInput type="date" value={f.pulloutDate} onChange={(v) => updateForm(name, "pulloutDate", v)} /></Field><Field label="Branch Closure Date"><TextInput type="date" value={f.closureDate} onChange={(v) => updateForm(name, "closureDate", v)} /></Field><Field label="Store Manager Name" required><TextInput value={f.storeManagerName} onChange={(v) => updateForm(name, "storeManagerName", v)} /></Field><Field label="Removed By" required><TextInput value={f.removedBy} onChange={(v) => updateForm(name, "removedBy", v)} /></Field><Field label="In-charge Name" required><TextInput value={f.inchargeName} onChange={(v) => updateForm(name, "inchargeName", v)} /></Field><Field label="Destination" required><TextInput value={f.destination} onChange={(v) => updateForm(name, "destination", v)} /></Field><Field label="Pullout Reason" required><TextInput value={f.pulloutReason} onChange={(v) => updateForm(name, "pulloutReason", v)} /></Field><Field label="Pullout Status"><SelectInput value={f.pulloutStatus} onChange={(v) => updateForm(name, "pulloutStatus", v)}>{["Pending Signature", "Submitted for Review", "Signed / Completed", "Returned for Correction"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field></> : <><Field label="Handover Date" required><TextInput type="date" value={f.handoverDate} onChange={(v) => updateForm(name, "handoverDate", v)} /></Field><Field label="Store Manager Name" required><TextInput value={f.storeManagerName} onChange={(v) => updateForm(name, "storeManagerName", v)} /></Field><Field label="Received By / Department"><TextInput value={f.receivedBy} onChange={(v) => updateForm(name, "receivedBy", v)} /></Field><Field label="Completed Person(s)" required><TextInput value={f.completedPersons} onChange={(v) => updateForm(name, "completedPersons", v)} /></Field><Field label="In-charge Name" required><TextInput value={f.inchargeName} onChange={(v) => updateForm(name, "inchargeName", v)} /></Field><Field label="Handover Status"><SelectInput value={f.handoverStatus} onChange={(v) => updateForm(name, "handoverStatus", v)}>{["Pending Signature", "Submitted for Review", "Signed / Completed", "Returned for Correction"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field></>}<div className="rounded-2xl bg-slate-50 p-4 md:col-span-2 xl:col-span-3"><p className="font-semibold">Devices Included Automatically</p><p className="text-sm text-slate-500">{assets.length} asset(s) will be included with serial numbers.</p></div>{isPullout && <label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Asset Condition</span><TextArea value={f.assetCondition} onChange={(v) => updateForm(name, "assetCondition", v)} /></label>}<label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">IT Remarks</span><TextArea value={f.itRemarks} onChange={(v) => updateForm(name, "itRemarks", v)} /></label><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Operations / Store Remarks</span><TextArea value={f.operationsRemarks} onChange={(v) => updateForm(name, "operationsRemarks", v)} /></label><ModalButtons save={isPullout ? "Save & Preview Pullout" : "Save & Preview Handover"} close={close} /></form></Modal>; }
function ProfileModal({ forms, updateForm, submit, formError, close, editId }) { const f = forms.profile; return <Modal title={editId ? "Edit User Profile" : "Add User / Profile"} subtitle="Create a personal portfolio and login account for each team member." onClose={close}>{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3"><Field label="Full Name" required><TextInput value={f.name} onChange={(v) => updateForm("profile", "name", v)} /></Field><Field label="Username" required><TextInput value={f.username} onChange={(v) => updateForm("profile", "username", v)} /></Field><Field label="Login Password"><TextInput type="password" value={f.password} onChange={(v) => updateForm("profile", "password", v)} placeholder={editId ? "Leave blank to keep current password" : "Default: ChangeMe@123 if blank"} /></Field><Field label="Email"><TextInput type="email" value={f.email || ""} onChange={(v) => updateForm("profile", "email", v)} placeholder="For password recovery verification" /></Field><Field label="Mobile"><TextInput value={f.mobile || ""} onChange={(v) => updateForm("profile", "mobile", v)} placeholder="For password recovery verification" /></Field><Field label="Access Level" required><SelectInput value={f.accessLevel} onChange={(v) => updateForm("profile", "accessLevel", v)}><option value="user">User - View Only</option><option value="admin">Admin - Full Access</option></SelectInput></Field><Field label="Role / Position" required><TextInput value={f.roleTitle} onChange={(v) => updateForm("profile", "roleTitle", v)} /></Field><Field label="Department" required><TextInput value={f.department} onChange={(v) => updateForm("profile", "department", v)} /></Field><Field label="Company" required><TextInput value={f.company} onChange={(v) => updateForm("profile", "company", v)} /></Field><Field label="Reporting To"><TextInput value={f.reportingTo} onChange={(v) => updateForm("profile", "reportingTo", v)} /></Field><Field label="Location / Branch Scope"><TextInput value={f.locationScope} onChange={(v) => updateForm("profile", "locationScope", v)} /></Field><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Specialization</span><TextArea value={f.specialization} onChange={(v) => updateForm("profile", "specialization", v)} /></label><label className="block md:col-span-2 xl:col-span-3"><span className="text-sm font-semibold text-slate-700">Profile Summary</span><TextArea value={f.profileSummary} onChange={(v) => updateForm("profile", "profileSummary", v)} /></label><ModalButtons save={editId ? "Update Profile" : "Create User & Profile"} close={close} /></form></Modal>; }
function CctvModal({ forms, updateForm, submit, formError, close }) {
  const f = forms.cctv;
  const reviewNames = ["Demo Technician", "Demo Engineer", "Siva", "Demo Team Lead", "Salahudeen", "Mohammed"];
  const autoDays = countRecordingDays(f.recordingStartDate, f.recordingEndDate);
  const statusFromDates = getAutoNvrStatus(autoDays);
  const updateRecordingDate = (field, value) => {
    const next = { ...f, [field]: value };
    const days = countRecordingDays(next.recordingStartDate, next.recordingEndDate);
    updateForm("cctv", field, value);
    updateForm("cctv", "recordingDays", days ? String(days) : "");
    updateForm("cctv", "reviewStatus", getAutoNvrStatus(days));
  };
  const updateReviewDate = (value) => {
    updateForm("cctv", "lastRecordingReviewDate", value);
    if (value) updateForm("cctv", "nextReviewDate", addDays(value, 7));
  };
  return <Modal title="Review / Update CCTV / NVR" subtitle="Enter recording start and end date. The app will calculate total recording days automatically." onClose={close} max="max-w-5xl">{formError && <ErrorBox text={formError} />}<form onSubmit={submit} className="mt-6 grid gap-4 md:grid-cols-2"><Field label="NVR Admin Username" required><TextInput value={f.nvrUsername} onChange={(v) => updateForm("cctv", "nvrUsername", v)} /></Field><Field label="NVR Password"><TextInput value={f.nvrPassword} onChange={(v) => updateForm("cctv", "nvrPassword", v)} /></Field><Field label="Recording Start Date" required><TextInput type="date" value={f.recordingStartDate} onChange={(v) => updateRecordingDate("recordingStartDate", v)} /></Field><Field label="Recording End Date" required><TextInput type="date" value={f.recordingEndDate} onChange={(v) => updateRecordingDate("recordingEndDate", v)} /></Field><div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><p className="text-sm font-semibold text-slate-700">Auto Calculated Recording Days</p><p className={`mt-2 text-3xl font-bold ${autoDays && autoDays < 60 ? "text-red-700" : autoDays ? "text-green-700" : "text-slate-500"}`}>{autoDays || 0} days</p><p className="mt-1 text-xs text-slate-500">Counted from start date to end date, including both dates.</p></div><div className="rounded-2xl border border-slate-200 bg-white p-4"><p className="text-sm font-semibold text-slate-700">Auto Status</p><div className="mt-2"><Badge type={!autoDays ? "warning" : autoDays < 60 ? "danger" : "success"}>{statusFromDates}</Badge></div><p className="mt-2 text-xs text-slate-500">60+ days is OK. Below 60 days will be marked for follow-up.</p></div><Field label="Last Recording Review Date" required><TextInput type="date" value={f.lastRecordingReviewDate} onChange={updateReviewDate} /></Field><Field label="Next Review Date"><TextInput type="date" value={f.nextReviewDate} onChange={(v) => updateForm("cctv", "nextReviewDate", v)} /></Field><Field label="Reviewed By" required><SelectInput value={f.reviewedBy} onChange={(v) => updateForm("cctv", "reviewedBy", v)}>{reviewNames.map((name) => <option key={name}>{name}</option>)}</SelectInput></Field><Field label="Review Status"><SelectInput value={f.reviewStatus || statusFromDates} onChange={(v) => updateForm("cctv", "reviewStatus", v)}>{["OK", "Below 60 Days", "Need Follow-up", "NVR Offline", "Camera Missing", "HDD Issue", "Password Pending", "Vendor Follow-up", "Review Pending"].map((x) => <option key={x}>{x}</option>)}</SelectInput></Field><div className="rounded-2xl bg-blue-50 p-4 text-sm text-blue-800 md:col-span-2"><p className="font-semibold">How to use this</p><p className="mt-1">In Hik-Connect playback calendar, check the oldest available recorded date and enter it as Recording Start Date. Keep Recording End Date as today, or the latest available date. Recording Days will calculate automatically.</p></div><label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Issue Found</span><TextArea value={f.issueFound} onChange={(v) => updateForm("cctv", "issueFound", v)} placeholder="Example: Recording starts from 2026-05-10 only / HDD storage low / camera channel missing" /></label><label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Action Required</span><TextArea value={f.actionRequired} onChange={(v) => updateForm("cctv", "actionRequired", v)} placeholder="Example: Recheck next week / increase HDD / coordinate vendor" /></label><label className="block md:col-span-2"><span className="text-sm font-semibold text-slate-700">Remarks</span><TextArea value={f.remarks} onChange={(v) => updateForm("cctv", "remarks", v)} /></label><ModalButtons save="Save Weekly Review" close={close} /></form><div className="mt-5 rounded-2xl bg-yellow-50 p-4 text-sm text-yellow-800">Passwords are sensitive. Keep access limited to Admin only and back up the database securely.</div></Modal>;
}
