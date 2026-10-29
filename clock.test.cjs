const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
function clock() {
 const html = fs.readFileSync(path.join(__dirname, '../templates/clock.html'), 'utf8');
 const initial = {minutes_per_level:1,current_level:1,schedule:['10 / 20','20 / 40'],active_players:['A','B','C'],eliminated_players:[],events:[],max_rebuys:1};
 const elements = new Map();
 function element() { return {textContent:'',value:'',children:[],classList:{add(){},remove(){},toggle(){}},
 appendChild(child){this.children.push(child); if(!this.value)this.value=child.value;},
 set innerHTML(value){this.children=[];this.value='';},play:()=>Promise.resolve(),pause(){}}; }
 let now=0,id=0; const timers=new Map();
 const context=vm.createContext({console,AbortController,Date:{now:()=>now},
 document:{getElementById(key){if(!elements.has(key))elements.set(key,element());return elements.get(key);},createElement:element,querySelectorAll:()=>[],addEventListener(){}},
 window:{addEventListener(){}},setTimeout(fn,delay){timers.set(++id,{fn,at:now+delay});return id;},clearTimeout(id){timers.delete(id);},
 fetch:async()=>({ok:true,json:async()=>({ok:true,state:{...initial,active_players:['A','B'],eliminated_players:['C']}})})});
 const source=html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('{{ texts | tojson }}',JSON.stringify({pause:'Pause',resume:'Resume'})).replace('{{ state | tojson }}',JSON.stringify(initial));
 vm.runInContext(source,context);
 return {context,timers,elements,run:code=>vm.runInContext(code,context),advance(ms){now+=ms;for(const[key,task]of[...timers]){if(task.at<=now){timers.delete(key);task.fn();}}}};
}
test('cancel kill resumes saved time',()=>{
 const c=clock();c.advance(10000);c.run('openKillModal()');c.advance(20000);c.run('closeKillModal()');c.advance(1000);
 assert.equal(c.run('running'),true);assert.equal(c.run('computeRemainingSeconds()'),49);assert.equal(c.timers.size,1);
});
test('elimination resumes running game',async()=>{
 const c=clock();c.run('openKillModal()');await c.run('submitKill()');c.advance(1000);
 assert.equal(c.run('running'),true);assert.equal(c.run('computeRemainingSeconds()'),59);assert.equal(c.timers.size,1);
});
test('manual pause survives cancellation and submission',async()=>{
 const c=clock();c.run('togglePause();openKillModal();closeKillModal()');assert.equal(c.run('running'),false);
 c.run('openKillModal()');await c.run('submitKill()');assert.equal(c.run('running'),false);assert.equal(c.timers.size,0);
});


test('rebuy uses server field names and resumes once',async()=>{
 const c=clock();let body;
 c.context.fetch=async(url,options)=>{body=JSON.parse(options.body);return {ok:true,json:async()=>({ok:true,state:{...c.run('state'),events:[{action:'Kill',killed:'A',rebuy:'Oui'}]}})};};
 c.run('openKillModal();openKillModal()');await c.run('submitKill()');
 assert.equal(body.killed,'A');assert.equal(c.run('rebuysForPlayer("A")'),1);assert.equal(c.timers.size,1);
});
test('network and validation errors permit cancel to restore running clock',async()=>{
 for(const response of [()=>Promise.reject(new Error('offline')),async()=>({ok:false,json:async()=>({ok:false,error:'invalid'})})]){
  const c=clock();c.context.fetch=response;c.run('openKillModal()');await c.run('submitKill()');
  assert.equal(c.run('running'),false);assert.equal(c.run('killPending'),false);
  c.run('closeKillModal()');assert.equal(c.run('running'),true);assert.equal(c.timers.size,1);
 }
});
test('double submissions make only one request',async()=>{
 const c=clock();let resolve;let calls=0;
 c.context.fetch=()=>{calls++;return new Promise(r=>resolve=r);};
 c.run('openKillModal()');const first=c.run('submitKill()');await c.run('submitKill()');
 c.run('closeKillModal()');assert.equal(c.run('killModalOpen'),true);
 resolve({ok:true,json:async()=>({ok:true,state:c.run('state')})});await first;
 assert.equal(calls,1);assert.equal(c.timers.size,1);
});
test('level request failure keeps scheduling ticks',async()=>{
 const c=clock();c.context.fetch=async()=>{throw new Error('offline');};
 await c.run('nextLevel()');assert.equal(c.timers.size,1);assert.equal(c.run('running'),true);
});
test('concurrent level requests are coalesced and unwrap state',async()=>{
 const c=clock();let resolve;let calls=0;
 c.context.fetch=()=>{calls++;return new Promise(r=>resolve=r);};
 const first=c.run('nextLevel()');await c.run('nextLevel()');
 resolve({ok:true,json:async()=>({ok:true,state:{...c.run('state'),current_level:2}})});await first;
 assert.equal(calls,1);assert.equal(c.run('state.current_level'),2);assert.equal(c.timers.size,1);
});
test('manual pause during a pending level request is preserved',async()=>{
 const c=clock();let resolve;c.context.fetch=()=>new Promise(r=>resolve=r);
 const first=c.run('nextLevel()');c.run('togglePause()');
 resolve({ok:true,json:async()=>({ok:true,state:{...c.run('state'),current_level:2}})});await first;
 assert.equal(c.run('running'),false);assert.equal(c.timers.size,0);
});


test('repeated open/close never creates duplicate timer loops',()=>{
 const c=clock();
 for(let i=0;i<20;i++)c.run('openKillModal();closeKillModal();closeKillModal()');
 assert.equal(c.timers.size,1);
 c.advance(1000);assert.equal(c.run('computeRemainingSeconds()'),59);
});
test('invalid JSON leaves kill dialog recoverable',async()=>{
 const c=clock();c.context.fetch=async()=>({ok:true,json:async()=>{throw new Error('invalid JSON');}});
 c.run('openKillModal()');await c.run('submitKill()');c.run('closeKillModal()');
 assert.equal(c.run('running'),true);assert.equal(c.timers.size,1);
});
test('kill dialog cannot race an in-flight level response',async()=>{
 const c=clock();let resolve;c.context.fetch=()=>new Promise(r=>resolve=r);
 const first=c.run('nextLevel()');c.run('openKillModal()');assert.equal(c.run('killModalOpen'),false);
 resolve({ok:true,json:async()=>({ok:true,state:{...c.run('state'),current_level:2}})});await first;
 c.run('openKillModal()');assert.equal(c.run('killModalOpen'),true);
});

test('a stalled kill times out and cancellation restores the clock',async()=>{
 const c=clock();
 c.context.fetch=(url,options)=>new Promise((resolve,reject)=>{
  options.signal.addEventListener('abort',()=>reject(Object.assign(new Error('timeout'),{name:'AbortError'})));
 });
 c.run('openKillModal()');const pending=c.run('submitKill()');c.advance(15000);await pending;
 assert.equal(c.run('killPending'),false);c.run('closeKillModal()');
 assert.equal(c.run('running'),true);assert.equal(c.timers.size,1);
});
test('final elimination stops every timer and cannot be resumed',async()=>{
 const c=clock();
 c.context.fetch=async(url)=>({ok:true,json:async()=>url==='/api/final'
  ? {winner:'A',pot:300,payouts:[{rank:1,player:'A',percent:100,amount:300}]}
  : {ok:true,state:{...c.run('state'),active_players:['A'],eliminated_players:['C','B']}}});
 c.run('openKillModal()');await c.run('submitKill()');c.run('togglePause()');
 assert.equal(c.run('running'),false);assert.equal(c.run('finished'),true);assert.equal(c.timers.size,0);
});
test('HTTP failure must not replace the valid game state',async()=>{
 const c=clock();c.context.fetch=async()=>({ok:false,json:async()=>({error:'unavailable'})});
 await c.run('nextLevel()');
 assert.equal(c.run('state.current_level'),1);assert.equal(c.timers.size,1);
});
test('level failure at zero waits before retrying',async()=>{
 const c=clock();let calls=0;c.context.fetch=async()=>{calls++;throw new Error('offline');};
 c.advance(60000);
 for(let i=0;i<10;i++)await Promise.resolve();
 assert.equal(calls,1);assert.equal(c.timers.size,1);
 c.advance(999);assert.equal(calls,1);
});
test('a paused game keeps its clock stopped after an invalid kill',async()=>{
 const c=clock();c.context.fetch=async()=>({ok:true,json:async()=>({ok:false,error:'invalid'})});
 c.run('togglePause();openKillModal()');await c.run('submitKill()');c.run('closeKillModal()');
 assert.equal(c.run('running'),false);assert.equal(c.timers.size,0);
});
