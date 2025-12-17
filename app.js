
function buildCalendar(){
 const m=document.getElementById("month").value;
 if(!m) return;
 const [y,mo]=m.split("-").map(Number);
 const days=new Date(y,mo,0).getDate();
 const cal=document.getElementById("calendar");
 cal.innerHTML="";
 for(let d=1;d<=days;d++){
  const date=new Date(y,mo-1,d);
  const wd=date.getDay();
  const isWeekend=wd==0||wd==6;
  const isWed=wd==3;
  const isThu=wd==4;

  const div=document.createElement("div");
  div.className="day"+(isWeekend?" weekend":"");
  div.innerHTML=`
  <strong>${d}/${mo}/${y}</strong> ${isWeekend?"(หยุด)":""}<br>
  <label><input type="checkbox" class="holiday" ${isWeekend?"checked":""}> วันหยุด</label>
  <label><input type="checkbox" class="half"> ลาครึ่งวัน</label><br>
  OPD <input type="number" class="opd" min="40" max="70" value="0">
  Check <input type="number" class="check" min="40" max="70" value="0">
  ADR <input type="number" class="adr" min="0" max="5" value="0">
  DIS <input type="number" class="dis" min="0" max="4" value="0"><br>
  ${isWed?`HIV <input type="number" class="hiv" min="16" max="20" value="0">`:``}
  ${isThu?`Asthma <input type="number" class="asthma" min="7" max="10" value="0">
           TB <input type="number" class="tb" min="1" max="3" value="0">`:``}
  `;
  cal.appendChild(div);
 }
}

function calculate(){
 const days=document.querySelectorAll(".day");
 let work=0, rawScore=0;
 days.forEach(d=>{
  const holiday=d.querySelector(".holiday").checked;
  const half=d.querySelector(".half").checked;
  if(!holiday){
   work+=half?0.5:1;
   let minutes=0;
   Object.keys(ACTIVITIES).forEach(k=>{
    const el=d.querySelector("."+k);
    if(el){
     rawScore+=el.value*ACTIVITIES[k].score;
     minutes+=el.value*ACTIVITIES[k].minute;
    }
   });
   if(minutes>420){
    const scale=420/minutes;
    Object.keys(ACTIVITIES).forEach(k=>{
     const el=d.querySelector("."+k);
     if(el) el.value=(el.value*scale).toFixed(2);
    });
   }
  }
 });
 const target=getTarget(work);
 const final=rawScore>target?target:rawScore;
 document.getElementById("summary").innerHTML=`
 <table class="table">
 <tr><th>ผู้ปฏิบัติงาน</th><td>${document.getElementById("staff").value||document.getElementById("staffCustom").value}</td></tr>
 <tr><th>หัวหน้างาน</th><td>${document.getElementById("chief").value||document.getElementById("chiefCustom").value}</td></tr>
 <tr><th>วันทำงาน</th><td>${work}</td></tr>
 <tr><th>คะแนนรวม</th><td>${final.toFixed(3)}</td></tr>
 </table>`;
}

function exportPDF(){ window.print(); }
