const DATA_URL = "content.json?v=3";
const LIBRARY_META = {
 python:{name:"Python · conceitos gerais",description:"Fundamentos da linguagem organizados para consulta rápida."},
 tkinter:{name:"Tkinter · biblioteca padrão",description:"Janelas, eventos, widgets e layouts da interface gráfica incluída na biblioteca padrão do Python."},
 json:{name:"json · biblioteca padrão",description:"Converta dados Python para JSON, leia arquivos, navegue em estruturas e trate erros de formato."},
 random:{name:"random · biblioteca padrão",description:"Sorteios, escolhas, embaralhamento e reprodutibilidade para simulações e testes."},
 sys:{name:"sys · biblioteca padrão",description:"Argumentos de execução, caminhos de importação, entrada e saída e informações do interpretador."},
 customtkinter:{name:"CustomTkinter · externa",description:"Widgets, argumentos, aparência e padrões de interface para criar aplicações modernas sobre Tkinter."},
 requests:{name:"Requests · externa",description:"Requisições HTTP, APIs, sessões, autenticação, arquivos, timeouts e tratamento de falhas."},
 win32com_client:{name:"win32com.client · pywin32",description:"Automação COM no Windows com exemplos de Excel, Word e Outlook, incluindo abertura, edição e encerramento seguro."},
 pandas_openpyxl:{name:"Pandas + OpenPyXL · externas",description:"Aprenda desde uma lista de dicionários até um arquivo Excel: escolher campos, renomear colunas, filtrar dados e formatar abas."}
};
const STOP = new Set(["como","que","para","por","com","sem","uma","um","o","a","os","as","do","da","de","no","na","em","meu","minha","fazer","usar","quero","tem","ter","e","ou"]);
const SYNONYMS = {
  "adicionar":["acrescentar","colocar","inserir","incluir"],"colocar":["adicionar","inserir","acrescentar"],
  "remover":["apagar","excluir","retirar","deletar"],"pegar":["acessar","obter","buscar"],
  "buscar":["procurar","encontrar","acessar"],"mexer":["usar","trabalhar","manipular"],
  "dicionario":["dict","chaves","valores"],"lista":["listas","colecao","colecoes"],
  "item":["valor","elemento"],"itens":["elementos","valores"],"texto":["string","strings"],
  "arquivo":["arquivos","ler","gravar","escrever"]
  ,"json":["serializar","serialização","json.loads","json.dumps","arquivo json"],
  "requests":["http","api","requisição","requisições","get","post"],
  "random":["sorteio","aleatório","embaralhar","randint","choice"],
  "sys":["argv","argumentos","interpretador","stdout","sys.path"],
  "win32com":["pywin32","excel","word","outlook","automação","com"],
  "cliente":["win32com.client","pywin32"],
  "pandas":["dataframe","planilha","excel","colunas","lista de dicionarios"],
  "openpyxl":["xlsx","formatar planilha","celulas","workbook","worksheet"]
};
const $ = (s, root=document) => root.querySelector(s);
const norm = s => String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^\p{L}\p{N}_]+/gu," ").trim();
const tokens = s => norm(s).split(/\s+/).filter(x=>x&&!STOP.has(x));
let chapters=[],flat=[],activeId="",searchTimer,activeLibrary="python";
const toc=$("#toc"),doc=$("#documentation"),results=$("#search-results"),welcome=$("#welcome");
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
function highlight(code){
 let x=escapeHtml(code);
 x=x.replace(/(#.*)$/gm,'<span class="tok-comment">$1</span>');
 x=x.replace(/("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*')/g,'<span class="tok-str">$1</span>');
 x=x.replace(/\b(\d+(?:\.\d+)?)\b/g,'<span class="tok-num">$1</span>');
 x=x.replace(/\b(def|return|if|elif|else|for|while|in|not|and|or|break|continue|True|False|None|with|as|try|except|pass|import|from)\b/g,'<span class="tok-key">$1</span>');
 x=x.replace(/\b(print|input|int|float|str|list|dict|len|range|enumerate|zip|type|isinstance|open)\b(?=\s*\()/g,'<span class="tok-fn">$1</span>');
 return x;
}
function addCode(code){
 const box=document.createElement("div");box.className="code-card";
 box.innerHTML='<button class="copy-button" type="button" aria-label="Copiar código">Copiar</button><pre><code>'+highlight(code)+'</code></pre>';
 $(".copy-button",box).addEventListener("click",async e=>{const b=e.currentTarget;try{await navigator.clipboard.writeText(code);b.textContent="Copiado";setTimeout(()=>b.textContent="Copiar",1300)}catch{b.textContent="Selecione o código";setTimeout(()=>b.textContent="Copiar",1500)}});
 return box;
}
const LABEL=/^(O que faz|Exemplo|Resultado|Como funciona|Se não usar|Uso comum(?: com arquivos)?|Sintaxe|Observação|Atenção|Importante|Relacionado|Argumentos(?: principais)?|Parâmetros|Como escolher|Regra|Dica|Retorna|Recebe|Use quando|Evite|Padrão):?/i;
function isCodeLine(s){return /(?:=|\(|\)|\[|\]|\{|\}|\.|\b(?:print|input|return|if|for|while|with|open|def|else|elif|in)\b|^\s*["'])/.test(s)||/^\s{2,}\S/.test(s)}
function renderReferenceLines(lines){
 const host=document.createElement("div");host.className="reference-list";
 lines.forEach(line=>{const parts=line.split(/\s{2,}/,2);if(parts.length<2){const p=document.createElement("p");p.className="prose-line";p.textContent=line;host.append(p);return}const row=document.createElement("div");row.className="reference-row";row.innerHTML='<code>'+highlight(parts[0].trim())+'</code><span>'+escapeHtml(parts[1].trim())+'</span>';host.append(row)});
 return host;
}
function renderLines(lines){
 const host=document.createElement("div");host.className="section-body";let i=0;
 while(i<lines.length){
  const line=lines[i],label=line.match(LABEL);
  if(label){
   const p=document.createElement("p");p.className="prose-line topic-label";
   const remainder=line.slice(label[0].length).trimStart();
   p.innerHTML="<strong>"+escapeHtml(label[1])+(line.slice(label[1].length).startsWith(":")?":":"")+"</strong>"+((label[1]==="Exemplo"&&remainder)?"":escapeHtml(line.slice(label[0].length)));host.append(p);i++;
   if(label[1]==="Exemplo"&&remainder){host.append(addCode(remainder));continue}
   if(/^Exemplo/i.test(line)&&i<lines.length){const group=[];while(i<lines.length&&!LABEL.test(lines[i])&&isCodeLine(lines[i]))group.push(lines[i++]);if(group.length)host.append(addCode(group.join("\n")))}
   continue;
  }
  if(i===0||isCodeLine(line)){const group=[];while(i<lines.length&&!LABEL.test(lines[i])&&(group.length===0||isCodeLine(lines[i])))group.push(lines[i++]);host.append(addCode(group.join("\n")));continue}
  const p=document.createElement("p");p.className="prose-line";p.textContent=line;host.append(p);i++;
 }
 return host;
}
function build(){
 flat=[];doc.replaceChildren();toc.replaceChildren();
 chapters.forEach(ch=>{
  const group=document.createElement("div");group.className="toc-group";
  const link=document.createElement("button");link.className="toc-chapter";link.type="button";link.setAttribute("aria-expanded","false");link.dataset.target=ch.id;
  link.innerHTML='<span class="toc-num">'+ch.number+'</span><span>'+escapeHtml(ch.title)+'</span><span class="toc-chevron" aria-hidden="true">⌄</span>';group.append(link);
  const childNav=document.createElement("div");childNav.className="toc-items";childNav.id="toc-items-"+ch.id;
  link.setAttribute("aria-controls",childNav.id);
  const node=document.createElement("section");node.className="chapter";node.id=ch.id;node.dataset.title=ch.title;
  const head=document.createElement("header");head.className="chapter-head";
  head.innerHTML='<div class="chapter-kicker">Capítulo '+ch.number+'</div><h2>'+escapeHtml(ch.title)+'</h2>';node.append(head);
  if(ch.number===12&&ch.intro&&ch.intro.length){const p=document.createElement("p");p.className="chapter-intro";p.textContent=ch.intro[0];node.append(p);node.append(renderReferenceLines(ch.intro.slice(1)))}
  else if(ch.intro&&ch.intro.length){const p=document.createElement("p");p.className="chapter-intro";p.textContent=ch.intro.join("\n");node.append(p)}
  if(ch.content&&ch.content.length){const cont=document.createElement("div");cont.className="chapter-content";cont.append(renderLines(ch.content));node.append(cont)}
  flat.push({id:ch.id,title:ch.title,path:"Capítulo "+ch.number,chapter:ch,section:null});
  ch.sections.forEach(s=>{
   const sl=document.createElement("a");sl.className="toc-link";sl.href="#"+s.id;sl.dataset.target=s.id;sl.textContent=s.number+" "+s.title;childNav.append(sl);
   const sec=document.createElement("section");sec.className="section";sec.id=s.id;sec.dataset.title=s.title;sec.dataset.chapter=ch.number;
   const h=document.createElement("h3");h.innerHTML="<span>"+escapeHtml(s.number)+"</span>"+escapeHtml(s.title);sec.append(h,renderLines(s.lines));node.append(sec);
   flat.push({id:s.id,title:s.title,path:"Capítulo "+ch.number+" · "+ch.title,chapter:ch,section:s});
  });
  if(ch.sections.length)group.classList.add("has-sections");group.append(childNav);toc.append(group);doc.append(node);
 });
 $("#chapter-count").textContent=chapters.length+" capítulos";
 $("#section-total").textContent=(flat.length-chapters.length)+" tópicos";
 flat.forEach((entry,index)=>{
  const target=document.getElementById(entry.id);if(!target)return;
  if(!entry.section&&entry.chapter.sections.length)return;
  const pager=document.createElement("nav");pager.className="section-pager";pager.setAttribute("aria-label","Navegação entre tópicos");
  const previous=flat[index-1],next=flat[index+1];
  pager.innerHTML=(previous?'<a class="pager-link previous" href="#'+previous.id+'"><small>ANTERIOR</small><span>← '+escapeHtml(previous.section?previous.section.title:previous.chapter.title)+'</span></a>':'<span></span>')+(next?'<a class="pager-link next" href="#'+next.id+'"><small>PRÓXIMO</small><span>'+escapeHtml(next.section?next.section.title:next.chapter.title)+' →</span></a>':'<span></span>');
  target.append(pager);
 });
 setupNavigation();
}
function setLibrary(id){
 const data={python:window.DOC_CONTENT,tkinter:window.TKINTER_CONTENT,customtkinter:window.CUSTOMTKINTER_CONTENT,json:window.JSON_CONTENT,random:window.RANDOM_CONTENT,sys:window.SYS_CONTENT,requests:window.REQUESTS_CONTENT,win32com_client:window.WIN32COM_CLIENT_CONTENT,pandas_openpyxl:window.PANDAS_OPENPYXL_CONTENT}[id];
 if(!data)return;
 activeLibrary=id;chapters=data.chapters;build();
 document.querySelectorAll(".library-button").forEach(button=>{
  const active=button.dataset.library===id;
  button.classList.toggle("active",active);button.setAttribute("aria-current",active?"page":"false");
 });
 const meta=LIBRARY_META[id];$("#library-name").textContent=meta.name;$("#welcome-description").textContent=meta.description+" Pesquise por argumento, método ou objetivo.";
 document.title=meta.name+" — Python de bolso";
 results.hidden=true;welcome.hidden=false;doc.hidden=false;$("#search-input").value="";closeMenu();
 window.scrollTo({top:0,behavior:"smooth"});
}
function setupNavigation(){
 document.querySelectorAll(".toc a").forEach(a=>a.addEventListener("click",closeMenu));
 document.querySelectorAll(".toc-chapter").forEach(button=>button.addEventListener("click",()=>{
  const group=button.closest(".toc-group");
  const expanded=group.classList.toggle("expanded");
  button.setAttribute("aria-expanded",String(expanded));
 }));
 const observer=new IntersectionObserver(entries=>{
  const best=entries.filter(e=>e.isIntersecting).sort((a,b)=>b.intersectionRatio-a.intersectionRatio)[0];if(!best)return;
  activeId=best.target.id;const f=flat.find(x=>x.id===activeId);if(f)$("#crumb-current").textContent=f.section?f.section.number+" "+f.section.title:"CAPÍTULO "+f.chapter.number;
  document.querySelectorAll(".toc a").forEach(a=>a.classList.toggle("active",a.dataset.target===activeId));
  const chapter=best.target.classList.contains("chapter")?best.target:best.target.closest(".chapter");
  document.querySelectorAll(".toc-group").forEach(g=>{
   const current=g.querySelector(".toc-chapter")?.dataset.target===chapter?.id;
   g.classList.toggle("expanded",current);
   g.querySelector(".toc-chapter")?.setAttribute("aria-expanded",String(current));
  });
 },{rootMargin:"-18% 0px -68% 0px",threshold:[0,.12,.35]});
 document.querySelectorAll(".chapter,.section").forEach(el=>observer.observe(el));
}
function searchScore(entry,q,ts){
 const ch=entry.chapter,s=entry.section;let score=0;
 const title=norm(entry.title),chapterTitle=norm(ch.title);
 const alias=[...(ch.aliases||[]),...(s?s.aliases||[]:[])].map(norm).join(" ");
 const syntax=norm(s?(s.lines||[]).slice(0,2).join(" "):"");
 const body=norm((s?(s.lines||[]).join(" "):"")+" "+(ch.content||[]).join(" "));
 if(title===q)score+=90;if(title.includes(q))score+=52;if(chapterTitle===q)score+=65;if(chapterTitle.includes(q))score+=43;
 if(alias.includes(q))score+=48;if(syntax.includes(q))score+=38;if(body.includes(q))score+=15;
 for(const t of ts){
  if(title.split(" ").includes(t))score+=18;else if(title.includes(t))score+=9;
  if(chapterTitle.split(" ").includes(t))score+=16;else if(chapterTitle.includes(t))score+=8;
  if(alias.includes(t))score+=12;if(syntax.includes(t))score+=10;if(body.includes(t))score+=Math.min(6,body.split(t).length-1)*2;
 }
 const expanded=new Set(ts.flatMap(t=>[t,...(SYNONYMS[t]||[])]));
 for(const t of expanded)if(!ts.includes(t)){if(title.includes(t))score+=8;if(chapterTitle.includes(t))score+=10;if(alias.includes(t))score+=14;if(syntax.includes(t))score+=5}
 if(q.includes("dicionari")&&ch.number===4)score+=entry.section?38:280;
 if((q.includes("adicionar")||q.includes("colocar")||q.includes("inserir"))&&q.includes("lista")&&s&&s.number==="3.4")score+=45;
 return score;
}
function snippet(entry){const lines=entry.section?(entry.section.lines||[]):(entry.chapter.intro||[]);const line=lines.find(x=>x.length>36)||lines[0]||entry.chapter.title;return escapeHtml(line.slice(0,170)+(line.length>170?"…":""))}
function doSearch(raw){
 const q=norm(raw),ts=tokens(raw);if(!q){results.hidden=true;welcome.hidden=false;doc.hidden=false;return}
 welcome.hidden=true;doc.hidden=true;results.hidden=false;$("#crumb-current").textContent="BUSCA";
 const ranked=flat.map(e=>({e:e,score:searchScore(e,q,ts)})).filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,12);
 results.replaceChildren();const head=document.createElement("div");head.className="results-head";
 head.innerHTML="<h2>Resultados da busca</h2><span>"+ranked.length+" resultado"+(ranked.length===1?"":"s")+"</span>";results.append(head);
 if(!ranked.length){const empty=document.createElement("div");empty.className="no-results";empty.textContent="Não encontrei esse termo. Tente buscar por uma palavra técnica, como append, items ou return.";results.append(empty);return}
 ranked.forEach(({e})=>{const a=document.createElement("a");a.className="result-card";a.href="#"+e.id;
  a.innerHTML='<div class="result-path">'+escapeHtml(e.path)+'</div><div class="result-title">'+(e.section?escapeHtml(e.section.number)+" — ":"")+escapeHtml(e.title)+'</div><div class="result-snippet">'+snippet(e)+'</div>';
  a.addEventListener("click",()=>{const targetId=e.id;setTimeout(()=>{$("#search-input").value="";results.hidden=true;welcome.hidden=false;doc.hidden=false;closeMenu();document.getElementById(targetId)?.scrollIntoView()},60)});results.append(a);
 });
}
function closeMenu(){document.body.classList.remove("menu-open")}
function toggleSearch(){const wrap=$(".search-wrap");if(window.matchMedia("(max-width:680px)").matches)wrap.classList.add("mobile-open");$("#search-input").focus()}
async function init(){
 try{
  let json=window.DOC_CONTENT;
  if(!json){const res=await fetch(DATA_URL);if(!res.ok)throw new Error("content.json could not be loaded");json=await res.json()}
  window.DOC_CONTENT=json;setLibrary("python");
  $("#library-switcher").addEventListener("click",e=>{const button=e.target.closest(".library-button");if(button)setLibrary(button.dataset.library)});
  const input=$("#search-input");
  input.addEventListener("input",()=>{clearTimeout(searchTimer);searchTimer=setTimeout(()=>doSearch(input.value),80)});
  input.addEventListener("keydown",e=>{if(e.key==="Escape"){input.value="";doSearch("");input.blur();$(".search-wrap").classList.remove("mobile-open")}if(e.key==="Enter")$(".result-card")?.click()});
  $("#welcome-search").addEventListener("click",toggleSearch);
  $("#welcome-search").addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" ")toggleSearch()});
  document.addEventListener("keydown",e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="k"){e.preventDefault();toggleSearch()}});
  $("#theme-toggle").addEventListener("click",()=>{const dark=document.documentElement.dataset.theme!=="dark";document.documentElement.dataset.theme=dark?"dark":"light";localStorage.setItem("python-pocket-theme",dark?"dark":"light")});
  const saved=localStorage.getItem("python-pocket-theme");if(saved)document.documentElement.dataset.theme=saved;
  $("#menu-toggle").addEventListener("click",()=>document.body.classList.toggle("menu-open"));$("#scrim").addEventListener("click",closeMenu);
  if(location.hash)requestAnimationFrame(()=>document.getElementById(decodeURIComponent(location.hash.slice(1)))?.scrollIntoView());
 }catch(err){doc.innerHTML='<p class="no-results">Não foi possível carregar os dados da documentação. Confira se content.js e content.json estão na mesma pasta do site.</p>';console.error(err)}
}
init();

