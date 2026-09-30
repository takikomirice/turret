import {adminEnvironment, AdminSheet, plain} from './admin-environment.mjs';

export function schedulingEnvironment() {
  const e=adminEnvironment({actualAutomation:true}), {c,sheets,triggers}=e;
  let now=Date.parse('2026-10-01T00:00:00Z'), serial=0;
  c.Date=class extends Date { constructor(...args){super(...(args.length?args:[now]));} static now(){return now;} };
  c.ScriptApp.getScriptId=()=> 'script-one';
  sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID'],['kid@example.com','生徒','100']]));
  const form={published:false,accepting:false,isPublished(){return this.published;},isAcceptingResponses(){return this.accepting;},
    setPublished(v){this.published=v;this.accepting=v;},setAcceptingResponses(v){this.accepting=v;},getPublishedUrl:()=> 'https://docs.google.com/forms/d/f/viewform'};
  c.FormApp={openById:()=>form};
  c.formMetadata_=()=>({settings:{collectsEmail:true}});
  c.Drive={Permissions:{list:()=>({permissions:[{type:'domain',role:'reader',view:'published',domain:'example.com'}]})}};
  const posts=new Map(), calls=[];
  const service={
    create(body,course){calls.push(['create',plain(body),course]);const post={...plain(body),id:'m'+(++serial),creationTime:new c.Date().toISOString(),updateTime:new c.Date().toISOString()};posts.set(post.id,post);return plain(post);},
    get(course,id){calls.push(['get',course,id]);if(!posts.has(id))throw Error('not found');return plain(posts.get(id));},
    patch(body,course,id,options){calls.push(['patch',plain(body),course,id,plain(options)]);const post=posts.get(id);for(const key of options.updateMask.split(',')){if(key in body)post[key]=body[key];else delete post[key];}post.updateTime=new c.Date().toISOString();return plain(post);},
    list(course,options){calls.push(['list',course,plain(options)]);return {courseWorkMaterial:[...posts.values()].map(plain)};},
    remove(course,id){calls.push(['remove',course,id]);posts.get(id).state='DELETED';}
  };
  c.Classroom={Courses:{CourseWorkMaterials:service}};
  c.ScriptApp.newTrigger=handler=>{const b={timeBased(){return b;},everyMinutes(n){assertInterval(n);return b;},create(){const id='t'+(++serial);const t={getUniqueId:()=>id,getHandlerFunction:()=>handler,getEventType:()=> 'CLOCK'};triggers.push(t);return t;}};return b;};
  function assertInterval(n){if(n!==5)throw Error('unexpected interval');}
  const record=c.saveManagedRecord_({id:'f1',kind:'form',courseId:'100',formId:'f',stage:'registered',materialTitle:'資料',title:'フォーム',className:'テスト',formUrl:form.getPublishedUrl(),input:{description:'回答してください',policy:{domains:['example.com'],emails:[]}}});
  const preview=(action,options={})=>c.previewManagedFormAction(record.id,action,options);
  const run=(action,options={})=>{const p=preview(action,options);return c.runManagedFormAction(record.id,action,p.fingerprint,true,options);};
  return {...e,form,posts,calls,service,record,preview,run,now:()=>now,advance:ms=>{now+=ms;},read:()=>c.loadManagedRecord_(record.id),
    fire:()=>{const t=triggers.find(t=>t.getHandlerFunction()==='managedFormScheduleTick_');c.managedFormScheduleTick_({triggerUid:t.getUniqueId()});}};
}
