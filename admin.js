import { initializeApp } from "https://www.gstatic.com/firebasejs/12.8.0/firebase-app.js";
import { getFirestore, collection, addDoc, getDocs, query, where, doc, getDoc, updateDoc, deleteDoc, writeBatch, orderBy, setDoc } from "https://www.gstatic.com/firebasejs/12.8.0/firebase-firestore.js";
import { getAuth, signOut } from "https://www.gstatic.com/firebasejs/12.8.0/firebase-auth.js";

const app = initializeApp({
  apiKey: "AIzaSyDyyTKs5JVvTScGO3vdfOJBX7DCjv-M8ZQ",
  authDomain: "madrasa-exam-result.firebaseapp.com",
  projectId: "madrasa-exam-result",
  storageBucket: "madrasa-exam-result.firebasestorage.app",
  messagingSenderId: "852087187745",
  appId: "1:852087187745:web:657b0a8ffb23b44663b660"
});

const db = getFirestore(app);
const auth = getAuth(app);
const mid = new URLSearchParams(window.location.search).get("mid");

let globalPassMark = 40; 
let isResultPublished = false;
let markStudentMap = {}; 
let currentMarkSubjects = []; 
let currentMarkStudentDocId = null;

const showLoader = (show, text="") => {
    document.getElementById("overlay").style.display = show ? "flex" : "none";
    if(document.getElementById("loadingText")) document.getElementById("loadingText").innerText = text;
};

function sortSubjects(arr) { arr.sort((a,b)=> (a.createdAt||0) - (b.createdAt||0)); }

window.onload = async () => {
    if(!mid) return document.body.innerHTML = "<h2 style='text-align:center; margin-top:50px;'>Invalid Link (Madrasa ID Missing)</h2>";
    const isAllowed = await loadSettings();
    if(!isAllowed) return; 
    await loadClasses();
    attachEventListeners();
};

function attachEventListeners() {
    document.getElementById("logoutBtn").onclick = logoutAdmin;
    document.getElementById("saveSettingsBtn").onclick = saveSettings;
    document.getElementById("visibilityBtn").onclick = toggleVisibility;
    document.getElementById("addClassBtn").onclick = addClass;
    document.getElementById("subClass").onchange = loadSubjects;
    document.getElementById("addSubjectBtn").onclick = addSubject;
    document.getElementById("stuClass").onchange = loadStudentsForClass;
    document.getElementById("addStuBtn").onclick = addStudent;
    document.getElementById("delAllStuBtn").onclick = deleteAllStudents;
    document.getElementById("markClass").onchange = loadMarkRegs;
    document.getElementById("markReg").onchange = loadStudentMarksForm;
    document.getElementById("delAllMarksBtn").onclick = deleteAllMarks;
    document.getElementById("attClass").onchange = loadAttendanceForm;
    document.getElementById("delAllAttBtn").onclick = deleteAllAttendance;
    document.getElementById("calcRankBtn").onclick = calculateAutoResultsBtn;
    document.getElementById("downloadPdfBtn").onclick = downloadPDF;
    document.getElementById("studentsExcel").addEventListener('change', handleStudentExcel);
    document.getElementById("marksExcel").addEventListener('change', handleMarksExcel);
    document.getElementById("downloadLabelsBtn").onclick = downloadDeskLabels;
}

async function logoutAdmin() {
    if(confirm("Are you sure you want to logout?")) {
        try { await signOut(auth); window.location.href = "index.html"; } 
        catch (error) { window.location.href = "index.html"; }
    }
}

async function loadSettings(){
  const d = await getDoc(doc(db, "madrasas", mid));
  if(d.exists()){
    const data = d.data();
    let isBlocked = false;
    if(data.status === 'blocked') isBlocked = true;
    if(data.expiryDate) {
        let today = new Date().toISOString().split("T")[0]; 
        if(today > data.expiryDate) isBlocked = true;
    }
    
    if(isBlocked) {
        document.body.innerHTML = `
            <div style="display:flex; height:100vh; align-items:center; justify-content:center; background:#f8d7da; color:#721c24; font-family:sans-serif; text-align:center; padding:20px;">
                <div>
                    <i class="fas fa-ban" style="font-size: 60px; color: #dc3545; margin-bottom: 20px;"></i>
                    <h1 style="font-size:28px; margin-bottom:10px;">Access Denied!</h1>
                    <p style="font-size:16px; line-height:1.6;">Your Madrasa's Admin Panel has been locked. <br>This might be because the subscription expired or it was blocked by the Super Admin.<br><br><b>Please contact the Developer to unlock (₹300/year).</b></p>
                </div>
            </div>`;
        return false; 
    }

    document.getElementById("mname").value = data.name || "";
    document.getElementById("mloc").value = data.location || "";
    document.getElementById("passmark").value = data.passmark || 40;
    globalPassMark = Number(data.passmark) || 40;
    document.getElementById("title").innerText = data.name || "New Madrasa Setup";
    isResultPublished = data.isPublished || false;
    updateVisibilityBtnUI();
    return true;
  } else {
    document.getElementById("title").innerText = "New Madrasa Setup";
    await setDoc(doc(db, "madrasas", mid), { name: "New Madrasa", passmark: 40, status: 'active', createdAt: Date.now() });
    return true;
  }
}

async function saveSettings() {
  let mn = document.getElementById("mname").value;
  let mp = document.getElementById("passmark").value;
  if(!mn || !mp) return alert("Fill Name and Pass Mark.");
  await updateDoc(doc(db,"madrasas",mid),{
    name: mn, location: document.getElementById("mloc").value, passmark: Number(mp)
  });
  globalPassMark = Number(mp);
  document.getElementById("title").innerText = mn;
  alert("Settings Saved!");
}

async function toggleVisibility() {
    isResultPublished = !isResultPublished;
    await updateDoc(doc(db,"madrasas",mid),{ isPublished: isResultPublished });
    updateVisibilityBtnUI();
    if(isResultPublished) alert("Result Published! (കുട്ടികൾക്ക് ഇപ്പോൾ റിസൾട്ട് കാണാം)");
    else alert("Result Locked! (കുട്ടികൾക്ക് ഇനിമുതൽ 'Coming Soon' എന്ന് കാണിക്കും)");
}

function updateVisibilityBtnUI() {
    const btn = document.getElementById("visibilityBtn");
    if(!btn) return;
    if(isResultPublished) {
        btn.innerHTML = "<i class='fas fa-lock'></i> Lock Result (Hide from Students)";
        btn.className = "danger";
    } else {
        btn.innerHTML = "<i class='fas fa-globe'></i> Publish Result (Show to Students)";
        btn.className = "success";
    }
}

async function loadClasses(){
    const q = query(collection(db, "madrasas", mid, "classes"), orderBy("createdAt"));
    const snap = await getDocs(q);
    const selects = document.querySelectorAll(".dynamic-class-select");
    let optionsHTML = "<option value=''>Select Class</option>";
    document.getElementById("classList").innerHTML = "";
    
    snap.forEach(d => {
        let cName = d.data().name;
        optionsHTML += `<option value="${d.id}">${cName}</option>`;
        document.getElementById("classList").innerHTML += `
        <div class="item">
            <b>${cName}</b>
            <div class="action-btns">
                <button class="icon-btn edit-btn" onclick="window.editClass('${d.id}', '${cName}')"><i class="fas fa-pen"></i></button>
                <button class="icon-btn delete-btn" onclick="window.delDoc('${d.id}','classes')"><i class="fas fa-trash"></i></button>
            </div>
        </div>`;
    });
    selects.forEach(s => { let currentVal = s.value; s.innerHTML = optionsHTML; if(currentVal) s.value = currentVal; });
}

async function addClass() {
    let cName = document.getElementById("className").value;
    if(!cName) return alert("Enter class name");
    showLoader(true);
    await addDoc(collection(db, "madrasas", mid, "classes"),{ name: cName, createdAt: Date.now() });
    document.getElementById("className").value = "";
    await loadClasses();
    showLoader(false);
}

window.editClass = async (id, oldName) => {
    let newName = prompt("Edit Class Name:", oldName);
    if(newName && newName !== oldName) {
        await updateDoc(doc(db, "madrasas", mid, "classes", id), { name: newName });
        loadClasses();
    }
};

async function loadSubjects(){
  const list = document.getElementById("subjectList");
  list.innerHTML="";
  const cId = document.getElementById("subClass").value;
  if(!cId) return;
  const q = query(collection(db, "madrasas", mid, "subjects"), where("class","==",cId));
  const snap = await getDocs(q);
  let subjects = [];
  snap.forEach(d => subjects.push({ id: d.id, ...d.data() }));
  sortSubjects(subjects); 

  subjects.forEach(s => {
    let max = s.maxMarks ? `(Max: ${s.maxMarks})` : '';
    list.innerHTML += `<div class="item"><span>${s.subject} <small style="color:#666">${max}</small></span><div class="action-btns"><button class="icon-btn delete-btn" onclick="window.delDoc('${s.id}','subjects')"><i class="fas fa-trash"></i></button></div></div>`;
  });
}

async function addSubject() {
  let cId = document.getElementById("subClass").value;
  let cSelect = document.getElementById("subClass");
  let cName = cSelect.options[cSelect.selectedIndex].text;
  let sName = document.getElementById("subName").value;
  let sMax = document.getElementById("subMax").value;
  if(!cId || !sName) return alert("Fill subject details");
  
  let data = { class: cId, className: cName, subject: sName, createdAt: Date.now() };
  if(sMax) data.maxMarks = Number(sMax); 
  await addDoc(collection(db, "madrasas", mid, "subjects"), data);
  document.getElementById("subName").value = ""; document.getElementById("subMax").value = "";
  loadSubjects();
}

let editingStudentId = null;
function loadStudentsForClass() {
    editingStudentId = null; 
    document.getElementById("addStuBtn").innerText = "Add";
    document.getElementById("reg").value = ""; 
    document.getElementById("sname").value = "";
    loadStudents();
}

async function loadStudents(){
  const list = document.getElementById("studentList");
  list.innerHTML="";
  const cId = document.getElementById("stuClass").value;
  if(!cId) return;
  const q = query(collection(db, "madrasas", mid, "students"), where("class","==",cId));
  const snap = await getDocs(q);
  let students = [];
  snap.forEach(d => students.push({ id: d.id, ...d.data() }));
  students.sort((a, b) => String(a.reg).localeCompare(String(b.reg), undefined, { numeric: true }));

  students.forEach(s => {
    list.innerHTML += `<div class="item"><span style="font-weight:500">${s.reg} - ${s.name}</span><div class="action-btns"><button class="icon-btn edit-btn" onclick="window.editStudent('${s.id}', '${s.reg}', '${s.name}')"><i class="fas fa-pen"></i></button><button class="icon-btn delete-btn" onclick="window.delDoc('${s.id}','students')"><i class="fas fa-trash"></i></button></div></div>`;
  });
}

window.editStudent = (id, r, n) => {
    editingStudentId = id; document.getElementById("reg").value = r; document.getElementById("sname").value = n;
    document.getElementById("addStuBtn").innerText = "Update"; 
};

async function addStudent() {
  let cId = document.getElementById("stuClass").value;
  let sel = document.getElementById("stuClass");
  let cName = sel.options[sel.selectedIndex].text;
  let r = document.getElementById("reg").value.trim();
  let n = document.getElementById("sname").value.trim();
  
  if(!cId || !r || !n) return alert("Fill all fields");
  showLoader(true);
  
  if(editingStudentId) {
      await updateDoc(doc(db, "madrasas", mid, "students", editingStudentId), { reg: r, name: n });
      editingStudentId = null; document.getElementById("addStuBtn").innerText = "Add";
  } else {
      const q = query(collection(db, "madrasas", mid, "students"), where("class","==",cId), where("reg","==",r));
      const snap = await getDocs(q);
      if(!snap.empty) {
          alert("Error: ഈ ക്ലാസ്സിൽ ഇതേ രജിസ്റ്റർ നമ്പറിൽ മറ്റൊരു കുട്ടി നിലവിലുണ്ട്!");
          showLoader(false); return;
      }
      await addDoc(collection(db, "madrasas", mid, "students"),{ class:cId, className: cName, reg:r, name:n });
  }
  document.getElementById("reg").value=""; document.getElementById("sname").value=""; 
  loadStudents();
  showLoader(false);
}

async function deleteAllStudents() {
    let cId = document.getElementById("stuClass").value;
    if(!cId) return alert("Select a class");
    if(!confirm("Are you sure you want to DELETE ALL students in this class?")) return;
    showLoader(true);
    const q = query(collection(db, "madrasas", mid, "students"), where("class", "==", cId));
    const snap = await getDocs(q);
    const batch = writeBatch(db);
    snap.docs.forEach(d => batch.delete(d.ref));
    await batch.commit(); loadStudents(); showLoader(false);
}

window.delDoc = async (id, col) => {
    if(confirm("Are you sure you want to delete this?")){
        await deleteDoc(doc(db, "madrasas", mid, col, id));
        if(col === 'classes') loadClasses();
        if(col === 'subjects') loadSubjects();
        if(col === 'students') loadStudents();
    }
};

async function loadMarkRegs() {
    let mClass = document.getElementById("markClass").value;
    let markReg = document.getElementById("markReg");
    markReg.innerHTML = "<option value=''>Select Reg</option>";
    document.getElementById("markInputsContainer").style.display = "none"; 
    document.getElementById("markListContainer").innerHTML = "";
    document.getElementById("markStuName").innerText = "";
    markStudentMap = {}; currentMarkSubjects = [];
    currentMarkStudentDocId = null;
    if(!mClass) return;

    let q = query(collection(db, "madrasas", mid, "students"), where("class","==",mClass));
    let students = [];
    const snap = await getDocs(q);
    snap.forEach(d => { students.push(d.data().reg); markStudentMap[d.data().reg] = d.data().name; });
    students.sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
    students.forEach(r => markReg.innerHTML += `<option value="${r}">${r}</option>`);

    q = query(collection(db, "madrasas", mid, "subjects"), where("class","==",mClass));
    let subs = [];
    (await getDocs(q)).forEach(d => subs.push({ id: d.id, ...d.data() }));
    sortSubjects(subs); currentMarkSubjects = subs; 
}

async function loadStudentMarksForm(){
    let mClass = document.getElementById("markClass").value;
    let mReg = document.getElementById("markReg").value;
    let markInputsContainer = document.getElementById("markInputsContainer");
    
    if(!mReg || !mClass) return;
    document.getElementById("markStuName").innerText = "Student: " + markStudentMap[mReg];
    markInputsContainer.innerHTML = "Loading..."; markInputsContainer.style.display = "block";
    document.getElementById("markListContainer").innerHTML = "";
    currentMarkStudentDocId = null;

    const stuQ = query(collection(db, "madrasas", mid, "students"), where("class","==",mClass), where("reg","==",mReg));
    const stuSnap = await getDocs(stuQ);
    let stuRank = "-"; let stuGrade = "-"; let isPromoted = false; let studentStatus = "Not Calculated";
    
    if(!stuSnap.empty) {
        currentMarkStudentDocId = stuSnap.docs[0].id;
        let sData = stuSnap.docs[0].data();
        stuRank = sData.rank || "-"; stuGrade = sData.grade || "-";
        isPromoted = sData.isPromoted || false; studentStatus = sData.resultStatus || "Not Calculated";
    }

    const q = query(collection(db, "madrasas", mid, "marks"), where("class","==",mClass), where("reg","==",mReg));
    const docs = await getDocs(q);
    let existingMarks = {};
    docs.forEach(d => { existingMarks[d.data().subject] = { id: d.id, mark: d.data().mark }; });

    if(currentMarkSubjects.length === 0){ markInputsContainer.innerHTML = `<p style="color:red;">No subjects found.</p>`; return; }

    let html = `<div class="subject-grid">`;
    let listHtml = ""; let totalMarks = 0; let hasFailed = false;

    currentMarkSubjects.forEach(sub => {
        let markVal = existingMarks[sub.subject] ? existingMarks[sub.subject].mark : '';
        let docId = existingMarks[sub.subject] ? existingMarks[sub.subject].id : '';
        let maxStr = sub.maxMarks ? `(Max:${sub.maxMarks})` : '';
        html += `<div class="subject-box"><label>${sub.subject} <small>${maxStr}</small></label><input type="text" class="bulk-mark-input" data-sub="${sub.subject}" data-docid="${docId}" value="${markVal}" placeholder="Mark/A"></div>`;
        if(existingMarks[sub.subject]) {
            listHtml += `<div class="item"><span>${sub.subject}</span> <b>${existingMarks[sub.subject].mark}</b></div>`;
            if(typeof existingMarks[sub.subject].mark === 'number') {
                totalMarks += existingMarks[sub.subject].mark;
                if(existingMarks[sub.subject].mark < globalPassMark) hasFailed = true;
            } else hasFailed = true; 
        }
    });
    
    html += `</div>
             <div style="margin: 15px 0; background: #fff3cd; padding: 12px; border-radius: 8px; border-left: 4px solid #ffc107;">
                <label style="cursor:pointer; font-weight:bold; color:#856404; font-size: 14px; display:flex; align-items:center; gap:8px;">
                    <input type="checkbox" id="forcePromote" style="width:18px; height:18px;" ${isPromoted ? 'checked' : ''}>
                    Force Pass / Promote (പാസ്സ് മാർക്ക് ഇല്ലെങ്കിലും വിജയിപ്പിക്കാൻ ഇവിടെ ടിക്ക് ചെയ്യുക)
                </label>
             </div>
             <button id="saveMarksBtnId">Save All Marks</button>`;
             
    markInputsContainer.innerHTML = html;
    document.getElementById("saveMarksBtnId").onclick = saveAllMarks;

    if(listHtml !== "") {
        let statusColor = (studentStatus === "PASSED") ? "#28a745" : (studentStatus === "PROMOTED" ? "#17a2b8" : "#dc3545");
        let gradeStr = (stuGrade && stuGrade !== "-") ? ` | <strong>Grade:</strong> <span style='color:#6f42c1'>${stuGrade}</span>` : "";
        document.getElementById("markListContainer").innerHTML = `
            <h4 style="border-bottom: 2px solid #eee; padding-bottom: 5px; margin-bottom:10px; color:#444;">Saved Marks:</h4>
            <div class="list" style="margin-top:0;">${listHtml}</div>
            <div class="summary-box">
                <strong>Total:</strong> ${totalMarks} <br>
                <strong>Status:</strong> <strong style='color:${statusColor}'>${studentStatus}</strong> <br>
                <strong>Rank:</strong> <span style="color:#007bff; font-weight:bold;">${stuRank}</span> ${gradeStr}
            </div>
        `;
    }
}

// ==== MAIN OPTIMIZATION: BUNDLING DATA FOR RESULT PAGE ====
async function processRankCalculation(classId) {
    const stuQ = query(collection(db, "madrasas", mid, "students"), where("class","==",classId));
    const stuSnap = await getDocs(stuQ);
    if(stuSnap.empty) return;

    // 1. Fetch & Store Subjects List in Class Document
    const subQ = query(collection(db, "madrasas", mid, "subjects"), where("class","==",classId));
    const subSnap = await getDocs(subQ);
    let totalMaxMarksPossible = 0;
    let subjectsData = [];
    subSnap.forEach(s => {
        subjectsData.push(s.data());
        totalMaxMarksPossible += s.data().maxMarks ? Number(s.data().maxMarks) : 100;
    });
    sortSubjects(subjectsData);
    let subjectNames = subjectsData.map(s => s.subject); 
    
    // Save to Class doc so result page can read it in 1 request
    await updateDoc(doc(db, "madrasas", mid, "classes", classId), { subjectList: subjectNames });

    // 2. Fetch all Marks
    const markQ = query(collection(db, "madrasas", mid, "marks"), where("class","==",classId));
    const markSnap = await getDocs(markQ);
    let tempMarksMap = {}; 
    markSnap.forEach(d => {
        const m = d.data();
        if(!tempMarksMap[m.reg]) tempMarksMap[m.reg] = {};
        tempMarksMap[m.reg][m.subject] = m.mark;
    });

    // 3. Fetch all Attendance
    const attQ = query(collection(db, "madrasas", mid, "attendance"), where("class","==",classId));
    const attSnap = await getDocs(attQ);
    let attMap = {};
    attSnap.forEach(d => { attMap[d.data().reg] = d.data().attendance; });

    let studentMarksMap = {};
    Object.keys(tempMarksMap).forEach(reg => {
        let total = 0; let hasFailed = false;
        Object.values(tempMarksMap[reg]).forEach(markVal => {
            if(typeof markVal === 'number') {
                total += markVal;
                if(markVal < globalPassMark) hasFailed = true;
            } else { hasFailed = true; }
        });
        studentMarksMap[reg] = { total: total, hasFailed: hasFailed, marksObj: tempMarksMap[reg] };
    });

    let resultsArray = [];
    stuSnap.forEach(d => {
        let data = d.data();
        let sData = studentMarksMap[data.reg] || { total: 0, hasFailed: true, marksObj: {} }; 
        let isPromoted = data.isPromoted || false;
        let studentAtt = attMap[data.reg] || "-";
        
        let status = sData.hasFailed ? "FAILED" : "PASSED";
        if (isPromoted) status = "PROMOTED"; 
        
        let grade = "-";
        if (status === "FAILED" || status === "PROMOTED") { grade = "D"; } 
        else if (status === "PASSED") {
            if (totalMaxMarksPossible > 0) {
                let percentage = (sData.total / totalMaxMarksPossible) * 100;
                if(percentage >= 90) grade = "A+"; else if(percentage >= 80) grade = "A";
                else if(percentage >= 70) grade = "B+"; else if(percentage >= 60) grade = "B";
                else if(percentage >= 50) grade = "C+"; else if(percentage >= 40) grade = "C";
                else grade = "D+"; 
            } else { grade = "D+"; }
        }
        
        resultsArray.push({ 
            docId: d.id, reg: data.reg, name: data.name, total: sData.total, status: status, grade: grade,
            marksMapData: sData.marksObj, // Embed Marks Map
            attendanceData: studentAtt    // Embed Attendance
        });
    });

    resultsArray.sort((a, b) => b.total - a.total);

    // 4. Update Student Documents with everything bundled
    const batch = writeBatch(db);
    resultsArray.forEach((res, index) => {
        res.rank = (res.status === "PASSED" || res.status === "PROMOTED") ? index + 1 : "-"; 
        batch.update(doc(db, "madrasas", mid, "students", res.docId), { 
            totalMarks: res.total, 
            resultStatus: res.status, 
            rank: res.rank, 
            grade: res.grade,
            marksMap: res.marksMapData,   // Saved directly inside student
            attendance: res.attendanceData // Saved directly inside student
        }); 
    });
    await batch.commit();
}

async function saveAllMarks() {
    showLoader(true);
    let mClass = document.getElementById("markClass").value;
    let sel = document.getElementById("markClass");
    let cName = sel.options[sel.selectedIndex].text;
    const inputs = document.querySelectorAll('.bulk-mark-input');
    
    let isPromoted = document.getElementById("forcePromote") ? document.getElementById("forcePromote").checked : false;
    const batch = writeBatch(db);
    
    if(currentMarkStudentDocId) {
        batch.update(doc(db, "madrasas", mid, "students", currentMarkStudentDocId), { isPromoted: isPromoted });
    }

    inputs.forEach(input => {
        let sub = input.getAttribute('data-sub'), docId = input.getAttribute('data-docid'), val = input.value.trim();
        if (val !== "") {
            let markToSave = isNaN(val) ? val.toUpperCase() : Number(val);
            if (docId) batch.update(doc(db, "madrasas", mid, "marks", docId), { mark: markToSave });
            else batch.set(doc(collection(db, "madrasas", mid, "marks")), { class: mClass, className: cName, reg: document.getElementById("markReg").value, subject: sub, mark: markToSave });
        } else if (val === "" && docId) batch.delete(doc(db, "madrasas", mid, "marks", docId));
    });
    await batch.commit();
    await processRankCalculation(mClass);
    showLoader(false); loadStudentMarksForm(); alert("Marks Saved & Grade/Rank Updated!");
}

async function deleteAllMarks() {
    let mClass = document.getElementById("markClass").value;
    if(!mClass) return alert("Select class");
    if(!confirm("DELETE ALL MARKS for this class?")) return;
    showLoader(true);
    const q = query(collection(db, "madrasas", mid, "marks"), where("class", "==", mClass));
    const snap = await getDocs(q);
    const batch = writeBatch(db);
    snap.docs.forEach(d => batch.delete(d.ref));
    await batch.commit();
    document.getElementById("markClass").dispatchEvent(new Event('change'));
    showLoader(false);
}

async function loadAttendanceForm() {
    let aClass = document.getElementById("attClass").value;
    let attInputsContainer = document.getElementById("attInputsContainer");
    if(!aClass) { attInputsContainer.style.display = "none"; return; }
    attInputsContainer.innerHTML = "Loading..."; attInputsContainer.style.display = "block";

    let q = query(collection(db, "madrasas", mid, "students"), where("class","==",aClass));
    const snap = await getDocs(q);
    let students = []; snap.forEach(d => students.push(d.data()));
    students.sort((a, b) => String(a.reg).localeCompare(String(b.reg), undefined, { numeric: true }));

    const attQ = query(collection(db, "madrasas", mid, "attendance"), where("class","==",aClass));
    const existingDocs = await getDocs(attQ);
    let existingAtt = {}; existingDocs.forEach(d => { existingAtt[d.data().reg] = { id: d.id, val: d.data().attendance }; });

    let html = `<div class="subject-grid">`;
    students.forEach(s => {
        let attVal = existingAtt[s.reg] ? existingAtt[s.reg].val : '';
        let docId = existingAtt[s.reg] ? existingAtt[s.reg].id : '';
        html += `<div class="subject-box"><label>${s.reg} - ${s.name}</label><input type="text" class="bulk-att-input" data-reg="${s.reg}" data-docid="${docId}" value="${attVal}"></div>`;
    });
    html += `</div><button id="saveAttBtnId">Save Attendance</button>`;
    attInputsContainer.innerHTML = html;
    document.getElementById("saveAttBtnId").onclick = saveAllAttendance;
}

async function saveAllAttendance() {
    showLoader(true);
    let aClass = document.getElementById("attClass").value;
    let sel = document.getElementById("attClass");
    let cName = sel.options[sel.selectedIndex].text;
    const inputs = document.querySelectorAll('.bulk-att-input');
    const batch = writeBatch(db);
    inputs.forEach(input => {
        let regNo = input.getAttribute('data-reg'), docId = input.getAttribute('data-docid'), val = input.value.trim();
        if (val !== "") {
            if (docId) batch.update(doc(db, "madrasas", mid, "attendance", docId), { attendance: val });
            else batch.set(doc(collection(db, "madrasas", mid, "attendance")), { class: aClass, className: cName, reg: regNo, attendance: val });
        } else if (val === "" && docId) batch.delete(doc(db, "madrasas", mid, "attendance", docId));
    });
    await batch.commit(); showLoader(false); alert("Attendance Saved!");
}

async function deleteAllAttendance() {
    let aClass = document.getElementById("attClass").value;
    if(!aClass) return alert("Select a class");
    if(!confirm("DELETE ALL ATTENDANCE for this class?")) return;
    showLoader(true);
    const q = query(collection(db, "madrasas", mid, "attendance"), where("class", "==", aClass));
    const snap = await getDocs(q);
    const batch = writeBatch(db);
    snap.docs.forEach(d => batch.delete(d.ref));
    await batch.commit(); document.getElementById("attClass").dispatchEvent(new Event('change')); showLoader(false);
}

async function calculateAutoResultsBtn() {
    const classId = document.getElementById("resClass").value;
    if(!classId) return alert("Select a class");
    showLoader(true);
    try {
        await processRankCalculation(classId);
        showLoader(false);
        alert("Rank calculated successfully! Download PDF now.");
    } catch (error) { showLoader(false); alert("Error calculating: " + error.message); }
}

async function downloadPDF() {
    const classId = document.getElementById("resClass").value;
    let sel = document.getElementById("resClass");
    const cNameText = sel.options[sel.selectedIndex].text;
    if(!classId) return alert("Select a class");
    
    showLoader(true);
    try {
        const subQ = query(collection(db, "madrasas", mid, "subjects"), where("class","==",classId));
        const subSnap = await getDocs(subQ);
        let subjectsData = [];
        subSnap.forEach(d => subjectsData.push(d.data()));
        sortSubjects(subjectsData);
        let subjects = subjectsData.map(s => s.subject); 

        const stuQ = query(collection(db, "madrasas", mid, "students"), where("class","==",classId));
        const stuSnap = await getDocs(stuQ);
        let students = [];
        stuSnap.forEach(d => { if(d.data().resultStatus) students.push(d.data()); });
        if(students.length === 0) { showLoader(false); return alert("Please click 'Calculate Rank' first."); }
        students.sort((a, b) => String(a.reg).localeCompare(String(b.reg), undefined, { numeric: true }));

        const markQ = query(collection(db, "madrasas", mid, "marks"), where("class","==",classId));
        const markSnap = await getDocs(markQ);
        let marksMap = {};
        markSnap.forEach(d => { let m = d.data(); if(!marksMap[m.reg]) marksMap[m.reg] = {}; marksMap[m.reg][m.subject] = m.mark; });

        const attQ = query(collection(db, "madrasas", mid, "attendance"), where("class","==",classId));
        const attSnap = await getDocs(attQ);
        let attMap = {}; attSnap.forEach(d => { attMap[d.data().reg] = d.data().attendance; });

        const { jsPDF } = window.jspdf; const doc = new jsPDF({ orientation: 'landscape' });
        let mName = document.getElementById("mname").value || "Madrasa Report";
        let mLoc = document.getElementById("mloc").value || "";
        let fullTitle = mLoc ? `${mName}, ${mLoc}` : mName;

        doc.setFontSize(16); doc.text(fullTitle, 14, 15);
        doc.setFontSize(12); doc.text(`Class: ${cNameText}  |  Pass Mark: ${globalPassMark}`, 14, 23);

        let tableHead = ['Reg No', 'Name'];
        subjects.forEach(sub => tableHead.push(sub.substring(0, 8))); 
        tableHead.push('Total', 'Rank', 'Grade', 'Att.', 'Status');

        let tableBody = [];
        students.forEach(s => {
            let row = [s.reg, s.name];
            subjects.forEach(sub => {
                let mark = (marksMap[s.reg] && marksMap[s.reg][sub] !== undefined) ? marksMap[s.reg][sub] : '-';
                row.push(mark);
            });
            row.push(s.totalMarks || 0); row.push(s.rank || '-'); row.push(s.grade || '-');
            row.push(attMap[s.reg] || '-'); row.push(s.resultStatus || '-');
            tableBody.push(row);
        });

        doc.autoTable({
            startY: 30, head: [tableHead], body: tableBody, theme: 'grid',
            headStyles: { fillColor: [41, 128, 185], fontSize: 10 }, bodyStyles: { fontSize: 9 },
            didParseCell: function (data) {
                let statusColIndex = tableHead.length - 1;
                if (data.column.index === statusColIndex) {
                    if (data.cell.raw === 'FAILED') { data.cell.styles.textColor = [220, 53, 69]; data.cell.styles.fontStyle = 'bold'; } 
                    else if (data.cell.raw === 'PROMOTED') { data.cell.styles.textColor = [23, 162, 184]; data.cell.styles.fontStyle = 'bold'; }
                }
            }
        });
        doc.save(`Full_Data_${cNameText}.pdf`);
    } catch (e) { alert("PDF Error: " + e.message); }
    showLoader(false);
}

async function handleStudentExcel(e) {
    let cId = document.getElementById("stuClass").value;
    let sel = document.getElementById("stuClass");
    if(!cId) { e.target.value = ""; return alert("Please select a class first!"); }
    
    let cName = sel.options[sel.selectedIndex].text;
    const file = e.target.files[0];
    if(!file) return;

    showLoader(true, "Reading Excel...");
    const reader = new FileReader();
    reader.onload = async (evt) => {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, {type: 'binary'});
            const firstSheet = workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
            
            if(rows.length === 0) { showLoader(false); return alert("Excel file is empty!"); }
            
            showLoader(true, `Processing ${rows.length} students...`);
            
            const existingQ = query(collection(db, "madrasas", mid, "students"), where("class","==",cId));
            const existingSnap = await getDocs(existingQ);
            let existingStudents = {};
            existingSnap.forEach(d => { existingStudents[String(d.data().reg).trim()] = d.id; });

            const batch = writeBatch(db);
            let count = 0;
            let updateCount = 0;

            for(let row of rows) {
                let rNo = row['RegNo'] || row['regno'] || row['Reg'] || row['reg'];
                let sName = row['Name'] || row['name'];
                
                if(rNo && sName) {
                    rNo = String(rNo).trim();
                    sName = String(sName).trim();
                    
                    if(existingStudents[rNo]) {
                        let docRef = doc(db, "madrasas", mid, "students", existingStudents[rNo]);
                        batch.update(docRef, { name: sName });
                        updateCount++;
                    } else {
                        let docRef = doc(collection(db, "madrasas", mid, "students"));
                        batch.set(docRef, { class:cId, className: cName, reg: rNo, name: sName });
                        count++;
                    }
                }
            }
            if(count > 0 || updateCount > 0) {
                await batch.commit();
                alert(`${count} പുതിയ കുട്ടികളെ ആഡ് ചെയ്തു, ${updateCount} കുട്ടികളെ അപ്ഡേറ്റ് ചെയ്തു!`);
                loadStudents();
            } else { alert("No valid columns found. Please ensure columns are 'RegNo' and 'Name'."); }
        } catch (err) { alert("Error reading Excel: " + err.message); }
        e.target.value = "";
        showLoader(false);
    };
    reader.readAsBinaryString(file);
}

async function handleMarksExcel(e) {
    let mClass = document.getElementById("markClass").value;
    let sel = document.getElementById("markClass");
    if(!mClass) { e.target.value = ""; return alert("Please select a class first!"); }
    let cName = sel.options[sel.selectedIndex].text;
    const file = e.target.files[0]; if(!file) return;

    showLoader(true, "Reading Marks Excel...");
    const reader = new FileReader();
    reader.onload = async (evt) => {
        try {
            const data = evt.target.result; const workbook = XLSX.read(data, {type: 'binary'});
            const firstSheet = workbook.SheetNames[0]; const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
            if(rows.length === 0) { showLoader(false); return alert("Excel file is empty!"); }
            showLoader(true, "Uploading Marks... This may take a moment.");
            
            const q = query(collection(db, "madrasas", mid, "subjects"), where("class","==",mClass));
            const subSnap = await getDocs(q);
            let validSubjects = []; subSnap.forEach(d => validSubjects.push(d.data().subject.toLowerCase().trim()));

            const existingMarksQ = query(collection(db, "madrasas", mid, "marks"), where("class","==",mClass));
            const existingMarksSnap = await getDocs(existingMarksQ);
            let existingMarksMap = {}; 
            existingMarksSnap.forEach(d => { let mData = d.data(); let mapKey = String(mData.reg).trim() + "_" + String(mData.subject).toLowerCase().trim(); existingMarksMap[mapKey] = d.id; });

            const batch = writeBatch(db); let count = 0;

            for(let row of rows) {
                let rNo = row['RegNo'] || row['regno'] || row['Reg'] || row['reg'];
                if(!rNo) continue; rNo = String(rNo).trim();

                Object.keys(row).forEach(key => {
                    let subKey = key.toLowerCase().trim();
                    if(validSubjects.includes(subKey)) {
                        let markVal = row[key];
                        if(markVal !== "" && markVal !== null && markVal !== undefined) {
                            let markToSave = isNaN(markVal) ? String(markVal).toUpperCase() : Number(markVal);
                            let actualSubName = subSnap.docs.find(d => d.data().subject.toLowerCase().trim() === subKey).data().subject;
                            let mapKey = rNo + "_" + subKey;
                            
                            if (existingMarksMap[mapKey]) {
                                let docRef = doc(db, "madrasas", mid, "marks", existingMarksMap[mapKey]);
                                batch.update(docRef, { mark: markToSave });
                            } else {
                                let docRef = doc(collection(db, "madrasas", mid, "marks"));
                                batch.set(docRef, { class: mClass, className: cName, reg: rNo, subject: actualSubName, mark: markToSave });
                            }
                            count++;
                        }
                    }
                });
            }
            if(count > 0) {
                await batch.commit(); await processRankCalculation(mClass); 
                alert(`${count} mark entries uploaded successfully! Rank updated.`);
                document.getElementById("markClass").dispatchEvent(new Event('change'));
            } else { alert("No valid marks found. Ensure column names match Subject names exactly."); }
        } catch (err) { alert("Error reading Excel: " + err.message); }
        e.target.value = ""; showLoader(false);
    };
    reader.readAsBinaryString(file);
}

// ==== DESK LABELS PDF GENERATION ====
async function downloadDeskLabels() {
    const classId = document.getElementById("resClass").value;
    let sel = document.getElementById("resClass");
    if(!classId) return alert("Please select a class first!");
    
    const cNameText = sel.options[sel.selectedIndex].text;
    showLoader(true, "Generating Desk Labels...");

    try {
        const stuQ = query(collection(db, "madrasas", mid, "students"), where("class","==",classId));
        const stuSnap = await getDocs(stuQ);
        let students = [];
        stuSnap.forEach(d => students.push(d.data()));

        if(students.length === 0) {
            showLoader(false);
            return alert("No students found in this class.");
        }

        students.sort((a, b) => String(a.reg).localeCompare(String(b.reg), undefined, { numeric: true }));

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

        const boxWidth = 62;
        const boxHeight = 38;
        const gapX = 5; 
        const gapY = 8; 

        // പേപ്പറിന്റെ നടുവിലായി വരുന്നതിന് കണക്കുകൂട്ടുന്ന പുതിയ കോഡ്
        const pageWidth = doc.internal.pageSize.getWidth(); 
        const totalContentWidth = (boxWidth * 3) + (gapX * 2);
        const marginX = (pageWidth - totalContentWidth) / 2; 
        const marginY = 15; 

        let col = 0; let row = 0;
        const maxCols = 3; 

        students.forEach((s, index) => {
            if (col >= maxCols) { col = 0; row++; }
            if (row >= 6) { doc.addPage(); col = 0; row = 0; }

            let x = marginX + (col * (boxWidth + gapX));
            let y = marginY + (row * (boxHeight + gapY));

            doc.setDrawColor(0);
            doc.setLineWidth(0.4);
            doc.rect(x, y, boxWidth, boxHeight);

            // 1. Name 
            doc.setTextColor(0, 0, 0);
            doc.setFontSize(11);
            doc.setFont("helvetica", "bold");
            doc.text("Name:", x + 4, y + 10);

            // കുട്ടിയുടെ പേര് 
            doc.setFont("helvetica", "normal");
            let nameLines = doc.splitTextToSize(s.name, boxWidth - 22);
            doc.text(nameLines, x + 18, y + 10);

            let nextY = y + 10 + ((nameLines.length - 1) * 5) + 9;

            // 2. Class
            doc.setFont("helvetica", "bold");
            doc.text("Class:", x + 4, nextY);
            doc.setFont("helvetica", "normal");
            doc.text(cNameText, x + 18, nextY);

            // 3. Reg. No
            nextY += 9;
            doc.setFont("helvetica", "bold");
            doc.text("Reg. No:", x + 4, nextY);
            doc.setFontSize(11);
            doc.setFont("helvetica", "normal");
            doc.text(String(s.reg), x + 23, nextY);

            col++;
        });

        doc.save(`Desk_Labels_${cNameText}.pdf`);
    } catch (e) {
        console.error(e);
        alert("Error generating PDF: " + e.message);
    }
    showLoader(false);
}