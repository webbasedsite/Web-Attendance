// ---------------------
// Config
// ---------------------
const webAppUrl = 'YOUR_WEB_APP_URL'; // Replace with your Google Script Web App URL
const statusEl = document.getElementById('status');
const historyEl = document.getElementById('history');

// ---------------------
// Service Worker (Optional for PWA offline)
// ---------------------
if('serviceWorker' in navigator){
  window.addEventListener('load', ()=>{
    navigator.serviceWorker.register('/sw.js').then(()=>console.log('SW registered'));
  });
}

// ---------------------
// Role & Login Handling
// ---------------------
window.onload = async function(){
  const role = localStorage.getItem('role');
  const phone = localStorage.getItem('phone');
  if(role === 'agent'){
    populateAgentDashboard();
  } else if(role === 'incharge' && phone){
    showInchargeProfile();
  }
}

// ---------------------
// Incharge Login
// ---------------------
async function loginIncharge(){
  const phone = document.getElementById('phone').value.trim();
  const password = document.getElementById('password').value.trim();
  if(!phone || !password){statusEl.textContent="Enter phone & password"; return;}
  statusEl.textContent = "Checking...";
  try{
    const res = await fetch(webAppUrl,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:`action=login&phone=${phone}&password=${password}`});
    const data = await res.json();
    if(data.success){
      localStorage.setItem('phone',phone);
      localStorage.setItem('role',data.role);
      window.location.href = 'attendance.html';
    } else statusEl.textContent = data.message;
  }catch{statusEl.textContent="Login failed";}
}

// ---------------------
// Agent Dashboard
// ---------------------
async function populateAgentDashboard(){
  const agentDiv = document.getElementById('agentInfo');
  // Fetch office list from Google Sheet
  try{
    const res = await fetch(webAppUrl,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:`action=getOffices`});
    const data = await res.json();
    const officeSelect = document.getElementById('officeSelect');
    data.offices.forEach(o=>{
      let opt = document.createElement('option');
      opt.value = o.name;
      opt.textContent = `${o.name} (${o.number})`;
      officeSelect.appendChild(opt);
    });
    agentDiv.textContent="Select your office and shift";
  }catch{agentDiv.textContent="Failed to load offices";}
}

// ---------------------
// Attendance Handling
// ---------------------
async function attemptAttendance(action){
  const employeeId = localStorage.getItem('phone') || document.getElementById('officeSelect').value;
  const shift = document.getElementById('shiftSelect').value;
  if(!employeeId || !shift){statusEl.textContent="Select all fields"; return;}
  statusEl.textContent="Tracking location...";
  try{
    const officeData = await fetchOfficeLocation(employeeId);
    navigator.geolocation.getCurrentPosition(async (pos)=>{
      const userLat = pos.coords.latitude;
      const userLng = pos.coords.longitude;
      const distance = getDistance(userLat,userLng,officeData.lat,officeData.lng);
      if(distance>100){statusEl.textContent=`Too far from office (${distance.toFixed(1)}m)`; return
