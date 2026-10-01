import {test} from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment,plain} from './helpers/admin-environment.mjs';

function setup() {
 const {c}=adminEnvironment(),profile=plain(c.gradeDefault_());
 profile.base.valueHeader='評価';profile.base.dateHeader='授業日';
 const data={classes:[{id:'a',name:'A',students:[{email:'a@x',name:'A',number:'1'},{email:'b@x',name:'B',number:'2'}]}],responses:[],issues:[]};
 const response=(id,email='a@x',date='2026-09-24')=>({id,courseId:'a',email,date,fingerprint:id,fields:{評価:id}});
 return {c,profile,data,response};
}

test('explicit unknown-date extra rule resolves confirmation and retains its value',()=>{
 const {c,profile,data,response}=setup();data.responses=[response('r1','a@x','')];
 profile.rules.dates=[{courseId:'a',from:'',to:'',mode:'extra'}];
 const result=plain(c.gradeBuild_(data,profile));
 assert.equal(result.issues.length,0);
 assert.deepEqual(result.classes[0].rows[0],['出席番号','名前','日付不明（予定外）']);
 assert.equal(result.classes[0].rows[1][2],'r1');
});

test('classwide lesson allocation is independent of response iteration order',()=>{
 const {c,profile,data,response}=setup();
 const single=response('single'),pair=[response('first','b@x'),response('second','b@x')];
 data.responses=[single,...pair];
 const duplicate=c.gradeBuild_(data,profile).issues.find(i=>i.type==='duplicate');
 profile.rules.duplicates=[{key:duplicate.key,fingerprint:duplicate.fingerprint,assignments:{first:1,second:2}}];
 const first=plain(c.gradeBuild_(data,profile));data.responses=[...pair,single];const second=plain(c.gradeBuild_(data,profile));
 const issues=result=>result.issues.map(i=>[i.type,i.email]).sort();
 assert.deepEqual(issues(first),issues(second));
 assert.ok(first.issues.some(i=>i.type==='duplicate'&&i.email==='a@x'),'single answer must also choose its lesson slot');
});

test('changing class calendar lesson count invalidates an earlier single-lesson duplicate choice',()=>{
 const {c,profile,data,response}=setup();profile.base.mapping.a='授業A';
 profile.calendar=[{id:'e1',title:'授業A',date:'2026-09-24',time:'09:00'}];
 data.responses=[response('first'),response('revision')];
 const duplicate=c.gradeBuild_(data,profile).issues.find(i=>i.type==='duplicate');
 profile.rules.duplicates=[{key:duplicate.key,fingerprint:duplicate.fingerprint,assignments:{first:0,revision:1}}];
 assert.equal(c.gradeBuild_(data,profile).issues.length,0);
 profile.calendar.push({id:'e2',title:'授業A',date:'2026-09-24',time:'10:00'});
 assert.ok(c.gradeBuild_(data,profile).issues.some(i=>i.type==='duplicate'),'calendar changed from one lesson to two; previous re-submission assumption needs review');
});

test('nested ICS alarm summary cannot replace the actual lesson title',()=>{
 const {c}=setup();const ics=['BEGIN:VCALENDAR','VERSION:2.0','BEGIN:VEVENT','UID:lesson','SUMMARY:授業A','DTSTART:20260924T090000','BEGIN:VALARM','ACTION:EMAIL','TRIGGER:-PT15M','SUMMARY:Reminder','DESCRIPTION:Reminder','ATTENDEE:mailto:test@example.invalid','END:VALARM','END:VEVENT','END:VCALENDAR'].join('\r\n');
 assert.equal(c.gradeParseIcs_(ics)[0].title,'授業A');
});

test('date and answer correction records require their respective identifying fields',()=>{
 const {c,profile}=setup();profile.rules.answers=[{courseId:'a',from:'2026-09-24',to:'2026-09-25',mode:'planned'}];
 assert.throws(()=>c.gradeValidate_(profile),/回答|形式|補正/);
});

test('changed answer never silently loses a saved individual correction',()=>{
 const {c,profile,data,response}=setup();data.responses=[response('r1')];profile.rules.answers=[{id:'r1',fingerprint:'r1',to:'2026-09-25',mode:'planned'}];
 assert.equal(c.gradeBuild_(data,profile).issues.length,0);data.responses[0].fingerprint='changed';
 assert.ok(c.gradeBuild_(data,profile).issues.some(i=>i.type==='date'&&i.stale));
});
test('removing the adopted response cannot silently revive a previously excluded response',()=>{
 const {c,profile,data,response}=setup();data.responses=[response('first'),response('revision')];const issue=c.gradeBuild_(data,profile).issues[0];profile.rules.duplicates=[{key:issue.key,fingerprint:issue.fingerprint,assignments:{first:0,revision:1}}];
 data.responses.pop();assert.ok(c.gradeBuild_(data,profile).issues.some(i=>i.type==='duplicate'));
});
