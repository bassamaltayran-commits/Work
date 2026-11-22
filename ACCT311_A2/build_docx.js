const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,HeadingLevel,BorderStyle,LevelFormat,PageBreak}=require('docx');
const F="Calibri";
const p=(t,o={})=>new Paragraph({spacing:{after:120,line:276},...o,children:(Array.isArray(t)?t:[t]).map(x=>typeof x==='string'?new TextRun({text:x,font:F,size:22}):x)});
const b=t=>new TextRun({text:t,bold:true,font:F,size:22});
const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});
const bl=t=>new Paragraph({numbering:{reference:"bul",level:0},spacing:{after:80},children:(Array.isArray(t)?t:[t]).map(x=>typeof x==='string'?new TextRun({text:x,font:F,size:22}):x)});
const W=9360;
function tbl(rows,widths,opts={}){
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:widths,rows:rows.map((r,i)=>new TableRow({tableHeader:i===0,children:r.map((c,j)=>{
    const head=i===0; const tot=opts.bold&&opts.bold.includes(i);
    return new TableCell({width:{size:widths[j],type:WidthType.DXA},margins:{top:60,bottom:60,left:100,right:100},
      shading:head?{type:ShadingType.CLEAR,fill:"1F3864",color:"auto"}:(tot?{type:ShadingType.CLEAR,fill:"E8EEF7",color:"auto"}:undefined),
      children:[new Paragraph({alignment:(j>0&&opts.right&&!head)?AlignmentType.RIGHT:AlignmentType.LEFT,children:[new TextRun({text:String(c),font:F,size:20,bold:head||tot,color:head?"FFFFFF":"000000"})]})]});
  })}))});
}
const sp=()=>new Paragraph({spacing:{after:60},children:[]});

const c=[];
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:"University of Prince Mugrin",font:F,size:24,bold:true})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:"ACCT 311 – Managerial Accounting for Hospitality",font:F,size:24})]}));
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:240},children:[new TextRun({text:"Assignment 2: Scenario-Based Hospitality Operating Budget and Management Recommendation",font:F,size:28,bold:true})]}));
c.push(tbl([["Item","Details"],["Student name","Bassam Khalid Eissa Altayran"],["Student ID","4411839"],["Section","[SECTION]"],["Instructor","Dr. Malik Mohammed"],["Case","Al Manar City Hotel – one-month operating budget"],["Submission date","[DATE]"]],[2800,6560]));
c.push(sp());

c.push(h1("1. Expected-Case Assumptions"));
c.push(tbl([["Assumption","Permitted range","Selected (Expected)","Brief justification"],
["Occupancy","62% – 76%","68%","Just below the 69% midpoint. A city hotel has steady weekday business demand, but weekends are usually softer, and I am pricing slightly above the midpoint, which can cost a little occupancy."],
["ADR","SAR 390 – 450","SAR 430","Slightly above the SAR 420 midpoint, reflecting a city-centre location and business guests who are less price-sensitive. Staying well below SAR 450 avoids losing leisure guests."],
["F&B covers","7,500 – 9,200","8,200","Slightly below the 8,350 midpoint, consistent with occupancy just below its midpoint: 2,856 occupied room-nights give a steady base of in-house diners, plus local walk-in and event guests."],
["Average F&B check","SAR 88 – 102","SAR 97","Slightly above the SAR 95 midpoint, in line with the hotel’s above-midpoint room pricing; assumes a normal meal mix with no major banquet premium."]],[1700,1500,1500,4660]));
c.push(sp());
c.push(p("Overall, I deliberately kept every expected-case assumption within a few points of the middle of its range. The expected case should describe the most likely month, not the best month, so I avoided both ends of each range and kept the rooms and F&B assumptions consistent with each other (higher occupancy normally supports higher covers)."));

c.push(h1("2. Expected-Case Rooms Forecast and Departmental Budget"));
c.push(tbl([["Line item","Formula / working","SAR"],
["Rooms available (room-nights)","140 rooms × 30 days","4,200 room-nights"],
["Occupied room-nights","4,200 × 68%","2,856 room-nights"],
["Rooms revenue","2,856 × SAR 430","1,228,080"],
["Less: Variable rooms expenses","2,856 × SAR 88","(251,328)"],
["Rooms contribution margin","1,228,080 − 251,328","976,752"],
["Less: Fixed departmental expenses","Given","(135,000)"],
["Rooms departmental profit","976,752 − 135,000","841,752"],
["RevPAR (supporting KPI)","1,228,080 ÷ 4,200  (or 68% × 430)","292.40"]],[3200,3660,2500],{right:true,bold:[7]}));

c.push(h1("3. Expected-Case F&B Forecast and Departmental Budget"));
c.push(tbl([["Line item","Formula / working","SAR"],
["Expected covers","Selected assumption","8,200 covers"],
["F&B revenue","8,200 × SAR 97","795,400"],
["Less: Variable F&B expenses","8,200 × SAR 54","(442,800)"],
["F&B contribution margin","795,400 − 442,800  (SAR 43 per cover)","352,600"],
["Less: Fixed departmental expenses","Given","(148,000)"],
["F&B departmental profit","352,600 − 148,000","204,600"]],[3200,3660,2500],{right:true,bold:[6]}));

c.push(h1("4. Expected-Case Total Departmental Income and GOP"));
c.push(tbl([["Line item","Formula / working","SAR"],
["Rooms departmental profit","Section 2","841,752"],
["F&B departmental profit","Section 3","204,600"],
["Total Departmental Income (TDI)","841,752 + 204,600","1,046,352"],
["Less: Undistributed operating expenses","Given","(315,000)"],
["Gross Operating Profit (GOP)","1,046,352 − 315,000","731,352"],
["Total revenue (supporting)","1,228,080 + 795,400","2,023,480"],
["GOP margin (supporting)","731,352 ÷ 2,023,480","36.1%"]],[3200,3660,2500],{right:true,bold:[3,5]}));
c.push(sp());
c.push(p("GOP = Total Departmental Income − Undistributed Operating Expenses, following the USALI summary operating statement structure covered in class."));

c.push(h1("5. Conservative Scenario – Changed Assumptions"));
c.push(p("I changed all four demand and revenue assumptions (the task requires at least two). The fixed and variable cost rates were kept as given."));
c.push(tbl([["Assumption","Expected","Conservative","Reason for change"],
["Occupancy","68%","63%","Reflects a softer month, e.g., fewer corporate bookings or new competing supply. Kept just above the 62% floor because the floor is a worst case, not a cautious case."],
["ADR","SAR 430","SAR 405","When demand weakens, hotels usually discount to protect occupancy, so lower occupancy and lower ADR tend to move together."],
["F&B covers","8,200","7,600","Fewer in-house guests means fewer breakfast and dinner covers."],
["Average check","SAR 97","SAR 91","Guests spend more cautiously and fewer premium/event orders are expected."]],[1700,1300,1400,4960]));

c.push(h1("6. Conservative Scenario – Recalculation"));
c.push(h2("Rooms"));
c.push(tbl([["Line item","Formula / working","SAR"],
["Occupied room-nights","4,200 × 63%","2,646 room-nights"],
["Rooms revenue","2,646 × SAR 405","1,071,630"],
["Less: Variable rooms expenses","2,646 × SAR 88","(232,848)"],
["Less: Fixed departmental expenses","Given","(135,000)"],
["Rooms departmental profit","1,071,630 − 232,848 − 135,000","703,782"]],[3200,3660,2500],{right:true,bold:[5]}));
c.push(h2("Food and Beverage"));
c.push(tbl([["Line item","Formula / working","SAR"],
["F&B revenue","7,600 × SAR 91","691,600"],
["Less: Variable F&B expenses","7,600 × SAR 54","(410,400)"],
["Less: Fixed departmental expenses","Given","(148,000)"],
["F&B departmental profit","691,600 − 410,400 − 148,000","133,200"]],[3200,3660,2500],{right:true,bold:[4]}));
c.push(h2("Total Departmental Income and GOP"));
c.push(tbl([["Line item","Formula / working","SAR"],
["Total Departmental Income","703,782 + 133,200","836,982"],
["Less: Undistributed operating expenses","Given","(315,000)"],
["Gross Operating Profit (GOP)","836,982 − 315,000","521,982"],
["GOP margin (supporting)","521,982 ÷ 1,763,230","29.6%"]],[3200,3660,2500],{right:true,bold:[3]}));

c.push(h1("7. Scenario Comparison and Sensitivity"));
c.push(h2("Required Scenario Comparison Table"));
c.push(tbl([["Measure","Expected Case","Conservative Case","Difference / Interpretation"],
["Occupancy","68%","63%","−5 points (−210 room-nights); the single largest driver of lost profit."],
["ADR","SAR 430","SAR 405","−SAR 25 (−5.8%); falls straight to profit because no cost is saved."],
["F&B covers","8,200","7,600","−600 covers (−7.3%); each cover lost costs SAR 43 of margin."],
["Average check","SAR 97","SAR 91","−SAR 6 (−6.2%); pure margin loss on every cover."],
["Rooms departmental profit","SAR 841,752","SAR 703,782","−SAR 137,970 (−16.4%); about 66% of the GOP decline."],
["F&B departmental profit","SAR 204,600","SAR 133,200","−SAR 71,400 (−34.9%); F&B is more sensitive in % terms due to its thin margin."],
["Total Departmental Income","SAR 1,046,352","SAR 836,982","−SAR 209,370 (−20.0%)."],
["GOP","SAR 731,352","SAR 521,982","−SAR 209,370 (−28.6%); undistributed costs are fixed, so the whole decline reaches GOP."]],[2100,1700,1700,3860]));
c.push(h2("Isolated effect of each assumption on GOP"));
c.push(p("To identify the main drivers, each assumption was changed on its own while the other three were held at the expected-case value:"));
c.push(tbl([["Assumption changed alone","Working","GOP impact (SAR)","Rank"],
["Occupancy 68% → 63%","210 room-nights × (430 − 88)","−71,820","1"],
["ADR SAR 430 → 405","2,856 room-nights × SAR 25","−71,400","2"],
["Average check SAR 97 → 91","8,200 covers × SAR 6","−49,200","3"],
["F&B covers 8,200 → 7,600","600 covers × (97 − 54)","−25,800","4"],
["Sum of isolated effects","","−218,220",""],
["Interaction effect (overlap)","210 × 25 + 600 × 6","+8,850",""],
["Total GOP change","","−209,370",""]],[3000,3260,1800,1300],{right:true,bold:[7]}));
c.push(sp());
c.push(p([b("Greatest effects on GOP: occupancy and ADR. "),"Occupancy has the largest effect because every lost room-night removes SAR 342 of contribution (ADR SAR 430 less SAR 88 variable cost), a contribution margin ratio of about 80%. One percentage point of occupancy equals 42 room-nights, or roughly SAR 14,364 of GOP. ADR is a very close second because a price reduction saves no cost at all: every riyal discounted is lost directly from profit, and it applies to all 2,856 room-nights (about SAR 2,856 of GOP per SAR 1 of ADR). By comparison, F&B has a much lower contribution margin (SAR 43 per cover, about 44%), so changes in covers hurt less, although the average check still matters because, like ADR, it falls straight to profit."]));
c.push(p("The comparison also shows operating leverage. Fixed costs total SAR 598,000 (135,000 + 148,000 + 315,000). Because these do not fall when demand falls, a 12.9% drop in total revenue (SAR 260,250) produces a 28.6% drop in GOP. The Rooms department is the hotel’s profit engine, contributing over 80% of Total Departmental Income in both scenarios."));

c.push(h1("8. Management Recommendation"));
c.push(p([b("Recommendation: a hybrid approach. "),"I recommend that management adopt the expected case as the main revenue target, but set spending and staffing on the conservative case, with the gap acting as a flexible-budget buffer."]));
c.push(p("The expected case is built on mid-range, internally consistent assumptions, so it is a fair and motivating target for the sales and revenue teams. Using only the conservative case as the target could lead to under-selling, and managers might release discounted rates too early. On the other hand, the GOP gap of more than SAR 209,000 is too large to ignore. If management staffs and purchases for 68% occupancy and the month turns out closer to 63%, labour hours, food purchases and discretionary spending would be committed before revenue arrives."));
c.push(p("In practice, the hybrid plan means: (1) revenue targets of SAR 1,228,080 for Rooms and SAR 795,400 for F&B; (2) the core rota and food purchasing planned for the conservative volumes (2,646 room-nights and 7,600 covers), with on-call staff and short-lead purchasing added only when pickup confirms higher demand; and (3) discretionary undistributed spending, such as marketing campaigns or minor repairs, released in stages once mid-month results are known. Because the variable costs per room (SAR 88) and per cover (SAR 54) are known, the budget can be flexed to the actual volume at month-end, so managers are evaluated against the costs that should have been incurred at the level of activity actually achieved. Even in the conservative case the hotel still earns a GOP of SAR 521,982, so the downside is manageable and does not justify an overly defensive plan."));

c.push(h1("9. Indicators to Monitor During the Month"));
c.push(tbl([["Indicator","Budget benchmark","Action if actual moves away from budget"],
["1. Occupancy / on-the-books pickup (weekly)","68% (2,856 room-nights); about 95 room-nights per day","Below plan: target corporate and government accounts, launch short-stay packages, open OTA channels, and reduce on-call housekeeping hours. Above plan: close discounted rate categories and add on-call staff."],
["2. ADR and RevPAR","ADR SAR 430; RevPAR SAR 292.40","If ADR falls under SAR 405, review discounting: avoid deep cuts and use value-adds (breakfast, late check-out) instead. If occupancy is strong, raise rates on high-demand dates. Track RevPAR so higher occupancy is not bought with too much discounting."],
["3. F&B average check and cost per cover","Check SAR 97; variable cost SAR 54 per cover","If the check drops, train staff on upselling, promote set menus and beverage pairings. If cost per cover rises above SAR 54, review portion sizes, waste and supplier prices, and adjust purchasing to actual covers."]],[2300,2300,4760]));

c.push(new Paragraph({children:[new PageBreak()]}));
c.push(h1("Appendix: Decision Trail"));
c.push(tbl([["Assumption / Decision","Evidence or Reasoning Used","Why This Was Reasonable"],
["Expected occupancy 68%","Range 62–76% (midpoint 69%); steady weekday demand, softer weekends, slightly premium pricing.","Most likely outcome, not best case; leaves room for both upside and downside."],
["Expected ADR SAR 430","Slightly above the SAR 420 midpoint; city-centre business demand; ADR and occupancy trade off.","Pairs a modest rate premium with slightly lower occupancy instead of assuming both are high."],
["Expected covers 8,200","Midpoint 8,350; linked to 2,856 occupied room-nights plus outside guests.","Keeps F&B consistent with the rooms forecast."],
["Expected average check SAR 97","Slightly above the SAR 95 midpoint; matches the hotel’s room price positioning.","Consistent with ADR choice; no unusual banquet activity assumed."],
["Cost rates and fixed costs used as given","Case instructions.","Ensures comparability; only demand and price assumptions are uncertain."],
["Conservative: changed all four demand/revenue assumptions","Weaker demand usually lowers occupancy, rates, covers and spend together.","More realistic than changing one variable in isolation; exceeds the two-change minimum."],
["Conservative values not set at range minimums","63%, SAR 405, 7,600, SAR 91 are cautious but plausible.","A conservative case is a cautious plan, not a worst-case stress test."],
["Isolated sensitivity test","Changed one assumption at a time from the expected case.","Clearly shows which assumptions drive GOP most (occupancy, then ADR)."],
["Hybrid recommendation","GOP range SAR 521,982–731,352; known variable cost rates allow flexing.","Keeps targets ambitious while protecting profit if demand is weak."],
["Three monitoring indicators","Highest-sensitivity drivers from Section 7.","Management attention goes where the money is."]],[2700,3500,3160]));
c.push(sp());
c.push(h2("References"));
c.push(p("ACCT 311 course materials, University of Prince Mugrin: Lecture 2 – USALI Summary Operating Statement (Total Departmental Income → GOP); Managerial Accounting Applications in Hospitality."));

const doc=new Document({numbering:{config:[{reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
styles:{default:{document:{run:{font:F,size:22}},
heading1:{run:{font:F,size:28,bold:true,color:"1F3864"},paragraph:{spacing:{before:300,after:120},keepNext:true}},
heading2:{run:{font:F,size:24,bold:true,color:"2E5597"},paragraph:{spacing:{before:200,after:100},keepNext:true}}}},
sections:[{properties:{page:{margin:{top:1440,bottom:1440,left:1440,right:1440}}},children:c}]});
Packer.toBuffer(doc).then(buf=>fs.writeFileSync(process.argv[2],buf));
