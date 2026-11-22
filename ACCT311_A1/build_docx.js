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
c.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:240},children:[new TextRun({text:"Assignment 1: Ethical Use of Managerial Accounting Information in Hospitality",font:F,size:28,bold:true})]}));
c.push(tbl([["Item","Details"],["Student name","Bassam Khalid Eissa Altayran"],["Student ID","4411839"],["Instructor","Dr. Malik Mohammed"],["Selected issue","Manipulating budgets or forecasts to make departmental performance appear stronger"]],[2800,6560]));
c.push(sp());

c.push(h1("1. The Ethical Issue and Why It Matters"));
c.push(p("Budget manipulation happens when a manager deliberately builds a budget or forecast that does not reflect his or her honest expectation, so that later results look better than they really are. The most common form is budgetary slack: setting revenue targets lower, or cost targets higher, than the manager truly expects. Once the year starts, the same manager may also move costs between accounts or delay spending so that the department “beats budget”. The figures may be accurate in a narrow sense, but the comparison they invite is misleading."));
c.push(p("This matters in hospitality because hotels are run on monthly departmental budgets, department heads are judged on departmental profit against plan, and their bonuses are often tied to beating budgeted GOP or departmental profit. Hotel owners are also frequently separate from the management company, so they rely on the budget and monthly report to judge management. A biased budget weakens the owner’s main accountability tool."));

c.push(h1("2. A Realistic Hospitality Situation"));
c.push(p("A 220-room business hotel in Riyadh is preparing next year’s budget. The F&B Director’s annual bonus is 15% of salary if F&B departmental profit beats the budget. Last year the restaurant served about 9,000 covers per month at an average check of SAR 95, and the F&B variable cost was about SAR 55 per cover. Two large conferences are already confirmed for the coming spring."));
c.push(p("When submitting the budget, the director forecasts only 8,000 covers per month, ignores the conferences because they are “not certain”, and budgets food cost at 40% of revenue instead of the 36% actually achieved. Mid-year, after a quiet summer, the department falls behind. The director then reclassifies staff meals and some kitchen equipment repairs to Administrative & General and Property Operations & Maintenance, which are undistributed expenses, and delays a planned refrigeration service until January. On the monthly report, F&B shows a favourable variance and the director receives the full bonus. The hotel’s GOP has not improved at all; the costs have only been moved or postponed."));

c.push(h1("3. Managerial Accounting Information Involved"));
c.push(bl([b("Operating budget and forecast: "),"covers, average check and food-cost percentage used to set the F&B revenue and expense targets."]));
c.push(bl([b("Departmental profit and GOP (USALI): "),"the measure on which the director is evaluated, and the line that separates departmental costs from undistributed costs."]));
c.push(bl([b("Budget variances: "),"the favourable revenue and cost variances that are reported to the General Manager and the owner."]));
c.push(bl([b("Cost classification and allocation: "),"the decision about which department a cost is charged to."]));
c.push(bl([b("Performance measures linked to pay: "),"the bonus formula that rewards beating budget rather than achieving a realistic plan."]));

c.push(h1("4. Effects on Stakeholders"));
c.push(p([b("Owners and lenders. "),"The owner approved spending and staffing for a lower level of business than was expected, and may have made capital or financing decisions based on a weaker forecast. Paying a bonus for “beating” an easy target transfers money from the owner to the manager without any real improvement. Lenders monitoring GOP also receive a distorted picture."]));
c.push(p([b("Employees. "),"Because the budget assumed fewer covers, the restaurant was scheduled with fewer staff. When the conferences arrived, the existing team carried the extra workload without extra hours or recognition. Honest managers in other departments look worse in comparison, which may teach them to pad their own budgets next year."]));
c.push(p([b("Guests. "),"Under-staffing during busy periods leads to slow service, and the postponed refrigeration service creates a food-safety and quality risk. In hospitality the guest experience is the product, so this directly harms reviews and repeat business."]));
c.push(p([b("Other department managers and the wider organisation. "),"A&G and Maintenance now carry costs they did not cause and cannot control, so their variances turn unfavourable. The hotel also loses reliable information for pricing, staffing and next year’s budget."]));

c.push(h1("5. Comparison of Sources"));
c.push(p("Jensen (2001) argues that tying rewards to beating a budget target causes managers to lie when the target is set and to “game” results when it is measured; he describes budgeting as a system that pays people to lie. Merchant (1985) provides empirical evidence for part of this view: in his study of manufacturing managers, the tendency to create budgetary slack was higher when budgets were used rigidly for evaluation, and lower when managers participated in budgeting and superiors could detect slack. The IMA Statement of Ethical Professional Practice (IMA, 2017) approaches the same problem from a professional duty angle: management accountants must communicate information fairly and objectively and disclose all information that could reasonably influence users’ understanding."));
c.push(p("The sources agree that the problem is real and is driven by how budgets are used, not only by individual dishonesty. They differ in the solution. Jensen concludes that the link between budget targets and pay should be broken. Merchant suggests that better participation and information can reduce slack without abandoning budget-based control. The IMA adds the individual responsibility that the other two sources play down: even if the system encourages gaming, the professional must still refuse to present misleading information. Together they suggest a combined answer of system design plus personal integrity."));
c.push(tbl([["Source","Main Idea Used","How It Relates to the Issue","How It Differs / Adds to Other Sources"],
["Jensen (2001), Harvard Business Review","Linking pay to beating budget targets encourages managers to lowball targets and game results.","Explains the F&B Director’s incentive to understate covers and overstate costs.","Most critical of budgeting; recommends separating rewards from budget targets."],
["Merchant (1985), Accounting, Organizations and Society","Slack increases when budgets are used rigidly for evaluation; participation and superior information can reduce it.","Shows why the GM and owner need independent forecast data (e.g., confirmed conferences).","Empirical evidence; more moderate, keeps budgets but improves how they are used."],
["IMA (2017), Statement of Ethical Professional Practice","Duties of competence, confidentiality, integrity and credibility; information must be fair, objective and complete.","Reclassifying costs and hiding variances breaches the credibility and integrity standards.","Focuses on the individual’s professional duty rather than on system design."]],[1900,2600,2400,2460]));

c.push(h1("6. Applying ACCT 311 Concepts"));
c.push(p([b("USALI departmental profit and GOP. "),"Under the USALI structure, departmental profit is revenue minus the direct expenses of that department, while costs such as A&G and Property Operations & Maintenance are undistributed and deducted only at GOP level. This structure exists so that each department head is evaluated on costs he or she controls. In the case, reclassifying staff meals and repairs to undistributed accounts raises F&B departmental profit but leaves GOP unchanged. Seeing the problem through USALI makes the manipulation clear: the director is exploiting the boundary between departmental and undistributed costs, which is exactly the boundary that responsibility accounting relies on."]));
c.push(p([b("Static versus flexible budgets and variance analysis. "),"If actual covers were 9,000 against a budget of 8,000, a static comparison will almost automatically show favourable revenue and profit variances. A flexible budget restates expected costs at the actual 9,000 covers. Using the real food-cost ratio of 36%, flexible food cost would be SAR 307,800 (9,000 × SAR 95 × 36%), while the padded 40% assumption allows SAR 342,000. The SAR 34,200 gap is slack, not efficiency, and flexing the budget with realistic cost standards exposes it."]));
c.push(p([b("Cost behaviour and contribution margin. "),"Because F&B has a high variable cost per cover, understating covers also understates variable costs, so the budget appears internally consistent. Analysing the contribution margin per cover (about SAR 40) and the department’s fixed costs separately helps the reviewer test whether the cover forecast and the cost forecast make sense together."]));

c.push(h1("7. Recommended Controls"));
c.push(bl([b("Independent forecast review. "),"The Director of Finance and the Revenue Manager should build an independent volume forecast from historical data, the group-booking pace and the events calendar. Any departmental budget that differs by more than an agreed tolerance (e.g., 5%) must be explained in writing."]));
c.push(bl([b("Flexible-budget evaluation. "),"Monthly reports should show both static and flexible-budget variances, and departmental performance should be judged on cost per occupied room or per cover against standard, not only on total profit versus the original budget."]));
c.push(bl([b("Locked chart of accounts and reclassification approval. "),"Moving any expense between departments after posting should require Director of Finance approval and be disclosed in the monthly commentary, with a quarterly audit of journal entries between departmental and undistributed accounts."]));
c.push(bl([b("Balanced bonus design. "),"Bonuses should combine financial results (judged against flexible budget and prior-year trend) with guest satisfaction, food-safety audit scores and employee turnover, so that beating an easy target is not enough."]));
c.push(bl([b("Ethics training and a reporting channel. "),"Finance and department heads should receive annual training based on the IMA standards, and a confidential channel should allow staff to report pressure to manipulate figures."]));

c.push(h1("8. Final Recommendation and Defence"));
c.push(p("A responsible hospitality manager should submit the budget he or she honestly believes is most likely, explain the uncertainty openly (for example, by showing an expected and a conservative scenario), and report variances, both favourable and unfavourable, without moving costs to hide them. In the case above, the director should have included the confirmed conferences, used the actual 36% food-cost history, and kept staff meals and kitchen repairs in F&B."));
c.push(p("I believe this is the right position for two reasons. First, a budget only has value as a planning and control tool if it is honest; a padded budget leads to wrong staffing, purchasing and pricing, so the hotel loses money even while the report looks good. Second, the costs of gaming fall on people who cannot see the numbers: guests, employees and owners. However, I agree with Jensen (2001) and Merchant (1985) that individual honesty is not enough if the system rewards dishonesty. The hotel should therefore combine personal integrity, as required by the IMA standards, with controls that make gaming difficult and unrewarding. This keeps budgets useful for motivating and evaluating managers while protecting the reliability of the hotel’s information."));

c.push(new Paragraph({children:[new PageBreak()]}));
c.push(h1("References"));
const ref=(a,i,r)=>new Paragraph({spacing:{after:120},indent:{left:720,hanging:720},children:[new TextRun({text:a,font:F,size:22}),new TextRun({text:i,font:F,size:22,italics:true}),new TextRun({text:r,font:F,size:22})]});
c.push(ref("Institute of Management Accountants. (2017). ","IMA statement of ethical professional practice",". IMA."));
c.push(ref("Jensen, M. C. (2001). Corporate budgeting is broken—let’s fix it. ","Harvard Business Review, 79","(10), 94–101."));
c.push(ref("Merchant, K. A. (1985). Budgeting and the propensity to create budgetary slack. ","Accounting, Organizations and Society, 10","(2), 201–210."));
c.push(ref("University of Prince Mugrin. (2026). ","ACCT 311 lecture 2: USALI summary operating statement"," [Lecture slides]. ACCT 311 – Managerial Accounting for Hospitality."));

c.push(h1("Appendix: Decision Trail"));
c.push(tbl([["Decision / Choice You Made","Evidence or Source Used","Why You Chose It"],
["Chose budget/forecast manipulation as the issue","Assignment issue list; Jensen (2001); Merchant (1985)","It is common in hotels where bonuses are tied to departmental profit, and it is well researched."],
["Built an F&B case in a Riyadh business hotel","Course examples of departmental budgets; local hotel context","F&B has high variable costs and many cost lines that can be reclassified, so the issue is realistic."],
["Combined budget slack with cost reclassification","USALI departmental vs undistributed structure (Lecture 2)","Shows both stages of manipulation: when the target is set and when results are reported."],
["Selected Jensen, Merchant and the IMA Statement","HBR, peer-reviewed journal, professional body","Three different perspectives: practitioner critique, empirical research and professional ethics."],
["Applied USALI, flexible budgeting and contribution margin","ACCT 311 topics","They directly show how the manipulation works and how it can be detected."],
["Recommended a mix of system controls and personal integrity","Comparison of the three sources","The sources disagree on the solution; combining them covers both causes of the problem."]],[2800,3200,3360]));
const doc=new Document({numbering:{config:[{reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
styles:{default:{document:{run:{font:F,size:22}},
heading1:{run:{font:F,size:28,bold:true,color:"1F3864"},paragraph:{spacing:{before:300,after:120},keepNext:true}},
heading2:{run:{font:F,size:24,bold:true,color:"2E5597"},paragraph:{spacing:{before:200,after:100},keepNext:true}}}},
sections:[{properties:{page:{margin:{top:1440,bottom:1440,left:1440,right:1440}}},children:c}]});
Packer.toBuffer(doc).then(buf=>fs.writeFileSync(process.argv[2],buf));
