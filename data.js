const data = {
 asOf:'2026-11-25',headcount:400,
 people:['Member A','Member B','Member C','Member D','Member E','Member F','Member G','Member H'],
 pillars:[
 {name:'Culture',color:'#073BD1',icon:'sun',prior:61,history:[46,49,48,56,55,61],kpis:[['Council awareness',46,80,'%',0],['Active participation rate',18,30,'%',2],['Belonging score',3.9,4.5,' / 5',[0,2]]]},
 {name:'Community',color:'#248373',icon:'network',prior:71,history:[52,57,55,63,67,71],kpis:[['Knowledge-sharing reach',84,100,'',3],['Connections made',12,20,'',3],['Learning value',72,80,'%',3]]},
 {name:'Business Priorities',color:'#7561ac',icon:'layers',prior:83,history:[59,62,69,68,78,83],kpis:[['External reach',48,50,'',1],['Talent pipeline',9,10,'',1],['Employees upskilled',5,5,'',1]]}
 ],
 // Metrics: level, name, actual, target (null = tracking), unit, optional comparison.
 initiatives:[
 ['Youth Council Introduction Video',0,'Member D, Member I','2026-11-23','2026-11-29','2026-11-23',true,[
 ['Reach','Unique viewers',36,50,'%','144 of 400 employees · preview audience'],['Engagement','Watch-through rate',56,60,'%','Tentative target'],['Impact','Council awareness lift',16,30,' pp','30% baseline to 46%'],['Impact','Inbound ideas and join requests',11,null,'','+3 since prior reading'],['Time','On-time delivery',null,100,'%']]],
 ['SAP Youth Workshops',2,'Member G','2026-11-30','2026-12-06',null,true,[
 ['Reach','Students registered',60,null,'','+12 since pilot registration opened'],['Reach','Students reached',48,null,'','+48 from pilot session'],['Engagement','Student satisfaction',4.4,4.2,' / 5'],['Engagement','Employee facilitators',5,null,'','+2 since prior reading'],['Impact','Increased interest in SAP careers',68,70,'%'],['Impact','Students entering pipeline',9,null,'','+3 since prior reading'],['Time','On-time delivery',null,100,'%']]],
 ['Steps Challenge',0,'Member H, Member E, Member J','2026-11-02','2026-11-29','2026-11-02',false,[
 ['Reach','Participation rate',18,30,'%','72 of 400 employees'],['Reach','Teams formed',9,null,'','+2 teams'],['Engagement','Final-week retention',null,60,'%','Weekly logging rate: week 1 86% · week 2 64% · week 3 61%'],['Engagement','Steps logged',3.6,5,'M','Daily average 0.15M vs milestone pace 0.18M'],['Impact','Feel more connected',61,70,'%'],['Impact','Plan to keep the habit',52,60,'%'],['Time','On-time delivery',null,100,'%']]],
 ['Internal Interview Series',1,'Unassigned','2026-11-02','2026-11-30',null,false,[
 ['Reach','Episodes / leaders featured',0,1,'','Monthly publishing target'],['Reach','Unique viewers per episode',84,null,'','+14 on pilot preview'],['Engagement','Watch-through rate',53,60,'%'],['Engagement','Questions and guest nominations',8,null,'','+2 since prior reading'],['Impact','Learned something useful about careers',72,70,'%'],['Impact','Follow-up connections',12,null,'','+4 since pilot preview'],['Time','On-time delivery',null,100,'%']]]
 ],
 // Illustrative anonymized overlap keys enable distinct employee reach.
 reach:[Array.from({length:144},(_,i)=>i+1),[],Array.from({length:72},(_,i)=>i+117),Array.from({length:84},(_,i)=>i+129)],
 survey:[[4.3,40],[4.4,48],[4.1,50],[4.5,34]],
 tasks:[
 [0,'Agree the video brief',[3,0],3,5,5],[0,'Draft the introduction script',[3],5,8,9],[0,'Review the storyboard',[1],8,10,12],[0,'Record council introductions',[2,3],12,15,15],[0,'Edit the first cut',[1],16,19,20],[0,'Review captions and branding',[0],20,null,23],[0,'Publish the launch video',[3],24,29,null],
 [1,'Confirm university contacts',[6],2,4,4],[1,'Scope the student workshop',[6,5],4,7,9],[1,'Draft the invitation',[5],9,null,10],[1,'Prepare the facilitator guide',[2],10,13,16],[1,'Map the registration journey',[0],16,null,17],[1,'Review the pilot agenda',[6],18,null,19],[1,'Confirm KAU participation',[6],20,27,null],[1,'Follow up with KFUPM',[5],20,30,null],[1,'Book session space at SAP',[2,6],24,null,null],
 [2,'Launch Fit@SAP registration',[7,4],2,3,3],[2,'Review week-two retention',[4],12,19,null],[2,'Collect participant feedback',[7],20,29,null],[3,'Produce the first interview',[],null,null,null]
 ],
 checklist:['Select a guest','Confirm availability','Prepare questions','Record interview','Edit footage','Review captions','Publish episode']
}

// Metric explanations: what it is, how it is calculated ({X:a:b} = clickable input), data source.
// Refs: P:p = pillar score · K:p:j = pillar KPI · M:i:j = initiative metric · S:k = overview measure.
data.info = {
 pillars:[
  {def:'Strengthening SAP KSA’s culture: engagement, belonging, collaboration and a motivating workplace.'},
  {def:'An inclusive professional network where young staff share ideas, learn from each other and build relationships across SAP KSA.'},
  {def:'Young staff bringing fresh ideas and practical solutions to SAP’s business goals in KSA.'}
 ],
 kpis:[
  [{def:'Share of employees who say they know what the Youth Council does.',how:'Latest pulse-survey level of agreement with “I know what the Youth Council does”. The Intro Video’s {M:0:2} tracks the change from the pre-launch baseline.',src:'Quarterly pulse survey (Microsoft Forms)',from:['M:0:2']},
   {def:'Share of KSA employees actively taking part in at least one Culture activity.',how:'Equals {M:2:0}: participants ÷ KSA headcount (400).',src:'Fit@SAP registration export',from:['M:2:0']},
   {def:'How strongly employees feel they belong at SAP KSA.',how:'Average rating (1–5) on the council-wide pulse question “I feel a sense of belonging at SAP KSA”. Supporting signal: {M:2:4}.',src:'Quarterly pulse survey (Microsoft Forms)',from:[]}],
  [{def:'Employees who watched or joined knowledge-sharing content.',how:'Equals {M:3:1} from the Interview Series.',src:'Microsoft Stream analytics',from:['M:3:1']},
   {def:'New professional connections triggered by council activity: coffee chats and mentoring requests.',how:'Equals {M:3:5}.',src:'Council follow-up form',from:['M:3:5']},
   {def:'Share of viewers who learned something useful from colleagues.',how:'Equals {M:3:4}.',src:'Pulse survey',from:['M:3:4']}],
  [{def:'Students and external participants engaged through council programmes.',how:'Equals {M:1:1}.',src:'Workshop attendance sheets',from:['M:1:1']},
   {def:'Students who entered SAP’s hiring pipeline after a council activity.',how:'Equals {M:1:5}.',src:'HR follow-up',from:['M:1:5']},
   {def:'SAP employees who built facilitation or presenting skills through council activities.',how:'Equals {M:1:3}.',src:'Council log',from:['M:1:3']}]
 ],
 metrics:[
  [{def:'Share of KSA employees who opened the intro video.',how:'Unique viewers ÷ KSA headcount (400).',src:'Microsoft Stream analytics'},
   {def:'Average share of the video that viewers actually watched.',how:'Average watch time ÷ video length, across all views.',src:'Microsoft Stream analytics'},
   {def:'Change in council awareness since before the video launched.',how:'Awareness now minus the pre-launch baseline, in percentage points (30% → 46% = +16 pp).',src:'Pulse survey: baseline and follow-up'},
   {def:'Ideas submitted and requests to join, triggered by the video’s call to action.',how:'Count of council interest-form submissions since launch. Tracking only, no target.',src:'Council interest form'},
   {def:'Whether the initiative finished by its planned end date.',how:'100% if the last task closes by the planned finish, minus 10 pp for each week late. Scored when the initiative closes.',src:'Microsoft Planner'}],
  [{def:'Students who signed up for a workshop.',how:'Count of registration-form entries. Tracking only.',src:'Registration form'},
   {def:'Students who actually attended a workshop.',how:'Count of attendees on the workshop attendance sheets.',src:'Attendance sheets'},
   {def:'How students rated the workshop.',how:'Average rating (1–5) on the post-event survey.',src:'Post-event survey'},
   {def:'SAP employees who facilitated or presented at a workshop.',how:'Count of facilitators logged per workshop. Tracking only.',src:'Council log'},
   {def:'Share of students reporting more interest in SAP careers after the workshop.',how:'% answering Agree or Strongly agree on the post-event survey.',src:'Post-event survey'},
   {def:'Students who applied for internships or graduate roles, or shared CVs.',how:'Count from HR follow-up after each workshop. Tracking only.',src:'HR follow-up'},
   {def:'Whether the initiative finished by its planned end date.',how:'100% if the last task closes by the planned finish, minus 10 pp for each week late. Scored when the initiative closes.',src:'Microsoft Planner'}],
  [{def:'Share of KSA employees who joined the Steps Challenge.',how:'Registered participants ÷ KSA headcount (400).',src:'Fit@SAP export'},
   {def:'Teams registered for the challenge.',how:'Count of teams in the Fit@SAP challenge. Tracking only.',src:'Fit@SAP export'},
   {def:'Share of participants still logging steps in the final week.',how:'Participants logging in the final week ÷ participants registered. Available once the final week ends.',src:'Fit@SAP export'},
   {def:'Total steps logged by all participants, against the shared milestone.',how:'Sum of logged steps ÷ shared milestone (5M).',src:'Fit@SAP export'},
   {def:'Share of participants who feel more connected to colleagues.',how:'% answering Agree or Strongly agree on the post-challenge survey.',src:'Post-challenge survey'},
   {def:'Share of participants who plan to keep the walking habit.',how:'% answering Agree or Strongly agree on the post-challenge survey.',src:'Post-challenge survey'},
   {def:'Whether the initiative finished by its planned end date.',how:'100% if the last task closes by the planned finish, minus 10 pp for each week late. Scored when the initiative closes.',src:'Microsoft Planner'}],
  [{def:'Interview episodes published, each featuring a senior SAP employee.',how:'Count of episodes published this month; target 1 per month.',src:'Council log'},
   {def:'Employees who watched an episode.',how:'Average unique viewers across published episodes.',src:'Microsoft Stream analytics'},
   {def:'Average share of an episode that viewers actually watched.',how:'Average watch time ÷ episode length, across all views.',src:'Microsoft Stream analytics'},
   {def:'Questions sent in and guests nominated for future episodes.',how:'Count of comments with questions plus nomination-form entries. Tracking only.',src:'Comments and nomination form'},
   {def:'Share of viewers who learned something useful about careers at SAP.',how:'% answering Agree or Strongly agree to “I learned something useful about careers at SAP”.',src:'Pulse survey'},
   {def:'Coffee chats and mentoring requests that followed an episode.',how:'Count of follow-ups logged on the council form. Tracking only.',src:'Council follow-up form'},
   {def:'Whether the initiative finished by its planned end date.',how:'100% if the last task closes by the planned finish, minus 10 pp for each week late. Scored when the initiative closes.',src:'Microsoft Planner'}]
 ],
 stats:[
  {name:'Employees reached',def:'Distinct SAP KSA employees who took part in at least one council initiative.',how:'Unique employees across {M:0:0}, {M:2:0} and {M:3:1}. Someone who joined several initiatives is counted once.',src:'Participant lists: attendance, Forms and Fit@SAP exports'},
  {name:'Students reached',def:'Students engaged through council workshops.',how:'Equals {M:1:1}.',src:'Workshop attendance sheets'},
  {name:'Average satisfaction',def:'Average participant rating across council surveys.',how:'Each survey’s average rating weighted by its number of responses, e.g. {M:1:2}.',src:'Post-event and pulse surveys'},
  {name:'Initiatives on track',def:'Share of initiatives whose focal points report them on track.',how:'Initiatives marked On track ÷ initiatives shown.',src:'Microsoft Planner goal status'}
 ],
 tips:{'Total tasks':'All Planner tasks in the selected initiatives. A shared task counts once.','Open':'Tasks not yet marked complete in Planner.','Completed':'Tasks marked complete in Planner.','Late / overdue':'Completed after the due date, or still open past it.','On time':'Completed on or before the due date. Tasks without a due date are excluded.'}
};
