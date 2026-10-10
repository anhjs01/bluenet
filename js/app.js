import{CONFIG}from"./config.js";
import{getAll,getByKey,put,del}from"./db.js";
import{activeLot,listLots,selectLot,createLot,updateLot,deleteLot}from"./lots.js";
import{unitsForActiveLot,stats,saveUnit,removeUnit}from"./inventory.js";
import{scanIdentification,closeScanner}from"./scanner.js";
import{audioTest,cleanupAudio}from"./audio-test.js";
import{exportExcel,exportCsv,exportJson,readExcelFile,rowToUnit}from"./excel.js";
import{sync}from"./sync.js";

const root=document.querySelector("#modalRoot");
const $=s=>document.querySelector(s);

let editing=null;
let draft={};

const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({
  "&":"&amp;",
  "<":"&lt;",
  ">":"&gt;",
  '"':"&quot;"
}[c]));

async function refresh(){
  const l=await activeLot();
  const u=await unitsForActiveLot();
  const s=stats(u);

  $("#activeLotName").textContent=
    l?l.name+" · "+l.date:"Sin lote";

  $("#stats").innerHTML=[
    ["Total",s.total],
    ["Entregadas",s.delivered],
    ["Pendientes",s.total-s.delivered],
    ["Reparables",s.reparable],
    ["No reparables",s.nonrepairable],
    ["Listas para empacar",s.ready],
    ["Empacadas",s.packed]
  ].map(x=>'<div class="stat"><b>'+x[1]+"</b><span>"+x[0]+"</span></div>").join("");

  const latest=[...u]
    .filter(x=>x.updatedAt)
    .sort((a,b)=>new Date(b.updatedAt)-new Date(a.updatedAt))
    .slice(0,5);

  const rank=new Map(latest.map((x,i)=>[x.id,i]));

  const filter=($("#idFilter")?.value||"").trim().toLowerCase();

  const filtered=filter
    ?u.filter(x=>String(x.unitId||"").toLowerCase().includes(filter))
    :u;

  const matches=filter
    ?u.filter(x=>String(x.unitId||"").toLowerCase().includes(filter))
    :[];

  $("#idFilterResult").textContent=
    filter
      ?(matches.length
        ?matches.map(x=>"casilla "+x.sequence).join(" · ")
        :"ID no encontrado")
      :"";

  $("#inventoryBody").innerHTML=filtered.map(x=>{
    const rnk=rank.get(x.id);
    const dot=
      rnk===0
        ?"newest"
        :rnk!==undefined&&rnk<3
          ?"recent"
          :rnk!==undefined
            ?"old"
            :"none";

    return '<tr class="'+
      ((x.delivered===true||x.deliveryStatus==="Entregada")
        ?"state-delivered"
        :x.diagnosis==="No reparable"
        ?"state-no"
        :x.packaging==="Empacado"
          ?"state-packed"
          :x.packaging==="Listo para empacar"
            ?"state-ready"
            :"")+
      '">'+
      "<td>"+x.sequence+"</td>"+
      "<td>"+esc(x.ticket)+"</td>"+
      '<td><b>'+esc(x.unitId)+"</b></td>"+
      "<td>"+esc(x.manufacturer)+"</td>"+
      "<td>"+esc(x.lotDate||"")+"</td>"+
      "<td>"+esc((x.delivered===true||x.deliveryStatus==="Entregada")?(x.deliveryDate||""):"")+"</td>"+
      "<td>"+esc(x.warrantyUntil||"")+"</td>"+
      "<td>"+esc(x.diagnosis)+"</td>"+
      "<td>"+
        (x.repairable===true
          ?"Sí"
          :x.repairable===false
            ?"No"
            :"")+
      "</td>"+
      "<td>"+esc((x.repairs||[]).join(" + "))+"</td>"+
      "<td>"+esc((x.delivered===true||x.deliveryStatus==="Entregada")?(String(x.observations||"").toUpperCase().includes("ENTREGADO")?x.observations:["ENTREGADO",x.observations||""].filter(Boolean).join(" · ")):x.observations)+"</td>"+
      "<td>"+esc(x.packaging)+"</td>"+
      '<td class="update-cell">'+
        (x.updatedAt
          ?'<span class="update-badge"><span class="update-dot '+
            dot+
            '"></span>'+
            relativeTime(x.updatedAt)+
            "</span>"
          :'<span class="update-badge"><span class="update-dot none"></span>Sin fecha</span>')+
      "</td>"+
      '<td class="actions-cell">'+
        '<button data-edit="'+x.id+'">Editar</button>'+
        '<button class="danger" data-del="'+x.id+'">Eliminar</button>'+
      "</td>"+
      "</tr>";
  }).join("");
}

function relativeTime(value){
  const ms=Math.max(0,Date.now()-new Date(value).getTime());
  const sec=Math.floor(ms/1000);

  if(sec<60){
    return sec<=1
      ?"hace un segundo"
      :"hace "+sec+" segundos";
  }

  const min=Math.floor(sec/60);

  if(min<60){
    return min+" "+(min===1?"minuto":"minutos");
  }

  const h=Math.floor(min/60);

  if(h<24){
    return h+" "+(h===1?"hora":"horas");
  }

  const d=Math.floor(h/24);

  if(d<7){
    return d+" "+(d===1?"día":"días");
  }

  const w=Math.floor(d/7);

  if(w<5){
    return w+" "+(w===1?"semana":"semanas");
  }

  const mo=Math.floor(d/30);

  if(mo<12){
    return mo+" "+(mo===1?"mes":"meses");
  }

  const y=Math.floor(d/365);

  return y+" "+(y===1?"año":"años");
}

function close(){
  cleanupAudio();
  root.innerHTML="";
  editing=null;
}

/* ── Test de audio independiente ── */
function openAudioTest(withGuide=false){
  cleanupAudio();

  const guide = withGuide ? `
    <aside class="headset-guide">
      <h3>Guía técnica F · RJ9 / 4P4C</h3>
      <p class="muted">Guía de apoyo para diagnosticar diademas de 4 conductores. Los códigos F1–F4 son referencias de esta aplicación; no son una norma universal de la industria. La función real de cada hilo depende del modelo, la placa y el cable utilizado.</p>
      <h4>Etiquetas de la placa</h4>
      <div class="guide-table">
        <div><b>MIC+ / M+ / MIC / IN+</b><span>Entrada positiva del micrófono. Puede incluir polarización para micrófonos electret; no conectes aquí un hilo solo por su color.</span></div>
        <div><b>MIC− / M− / GND_MIC / SGND</b><span>Retorno o referencia del circuito de micrófono. No siempre equivale a la tierra del altavoz.</span></div>
        <div><b>SPK+ / EAR+ / EP+ / OUT+</b><span>Señal de audio hacia el auricular o altavoz.</span></div>
        <div><b>SPK− / EAR− / EP− / OUT−</b><span>Retorno de audio del auricular. En algunos circuitos puede ser una salida activa y no una tierra común.</span></div>
      </div>
      <h4>Configuraciones descritas en la guía</h4>
      <p>Cuenta los contactos del conector macho siempre desde la misma vista: mira de frente los contactos metálicos y coloca la pestaña hacia abajo. Numera de izquierda a derecha (pines 1, 2, 3 y 4). Si cambias la orientación o miras el extremo opuesto del cable, el orden visual puede parecer invertido.</p>
      <div class="guide-table pinout">
        <div><b>F1 · Estándar A</b><span>Secuencia de referencia guardada en el proyecto: 1 Negro (MIC+) · 2 Rojo (MIC−) · 3 Verde (SPK+) · 4 Blanco (SPK−). Es una asignación orientativa; comprueba la placa del modelo antes de soldar.</span></div>
        <div><b>F2 · Estándar B · Cisco IP 79xx</b><span>Referencia histórica del proyecto: 1 Verde (SPK+) · 2 Negro (MIC+) · 3 Rojo (MIC−) · 4 Blanco (SPK−). No se pudo confirmar en documentación oficial una asignación universal por colores para toda la serie Cisco 79xx; valida el modelo exacto.</span></div>
        <div><b>F3 · Estándar C · Plantronics / Poly</b><span>Referencia histórica del proyecto: 1 Negro (MIC+) · 2 Verde (SPK+) · 3 Blanco (SPK−) · 4 Rojo (MIC−). No asumir que aplica a todas las diademas Poly/Plantronics; puede variar según cable, adaptador y modelo.</span></div>
        <div><b>F4 · Estándar D · Jabra / GN Netcom</b><span><strong>Secuencia de colores corregida para tu referencia F4:</strong> 1 Negro · 2 Rojo · 3 Verde · 4 Blanco. La función MIC/SPK de cada conductor debe confirmarse en el diagrama o en las pistas de la placa; la secuencia por sí sola no prueba una asignación universal de Jabra/GN.</span></div>
      </div>
      <h4>Cómo comprobar la configuración con seguridad</h4>
      <ol>
        <li><b>Desconecta la diadema</b> de la caja USB, teléfono y cualquier alimentación antes de medir o soldar.</li>
        <li><b>Identifica los pines</b> con la orientación indicada arriba y anota el color que llega a cada contacto.</li>
        <li><b>Usa continuidad</b> en un multímetro para seguir cada conductor desde el contacto hasta su punto de soldadura. No midas continuidad con el equipo alimentado.</li>
        <li><b>Busca las marcas de la placa</b> (MIC+, MIC−, SPK+, SPK−) o el esquema del modelo. Si no hay marcas ni esquema, no deduzcas la función por color: compara con una unidad idéntica que funcione.</li>
        <li><b>Prueba al final</b> el audio de salida y el micrófono por separado. Cambia un solo conductor por vez y registra la configuración que haya sido verificada.</li>
      </ol>
      <h4>Diagnóstico orientativo</h4>
      <ul>
        <li><b>Audio se escucha, micrófono mudo:</b> revisa continuidad del micrófono, MIC+ / MIC−, soldaduras y compatibilidad del adaptador. No siempre es una inversión de polaridad.</li>
        <li><b>Zumbido o estática:</b> revisa hilos pelados, soldaduras, blindaje y posibles contactos entre circuitos; no unas MIC− y SPK− sin esquema.</li>
        <li><b>Audio débil o ausente:</b> comprueba continuidad de SPK+/SPK−, el altavoz y que el adaptador corresponda al teléfono o caja USB.</li>
        <li><b>La diadema funciona en un equipo y en otro no:</b> puede tratarse de una diferencia de pinout o compatibilidad, no necesariamente de una avería.</li>
      </ul>
      <p class="guide-warning">Importante: las fuentes oficiales consultadas describen conectores y compatibilidad por modelo, pero no establecen un código de colores F1–F4 universal. Usa estas secuencias como referencias de trabajo, no como sustituto del esquema del modelo.</p>
      <p class="muted">Fuentes para contrastar modelos: <a href="https://www.cisco.com/en/US/docs/voice_ip_comm/cuipph/7960g_7940g/4_2_3/english/administration/guide/7900spc8.html" target="_blank" rel="noopener">manual Cisco 7940G/7960G (conector RJ‑9 de 4 conductores)</a> · <a href="https://www.jabra.com/en-gb/supportpages/jabra-link-14201-10/14201-10/faq/d568a393-4f46-49f5-a839-a8eec498d411" target="_blank" rel="noopener">guía Jabra de cableado por modelo</a>.</p>
    </aside>` : '';

  root.innerHTML=
    '<div class="modalback">'+
      '<div class="modal audio-test-modal '+(withGuide?'audio-guide-modal':'')+'">'+
        '<div class="audio-test-header">'+
          '<div>'+
            '<span class="audio-test-kicker">PRUEBA DE AUDIO</span>'+
            '<h2>'+(withGuide?'Test de audio con guía técnica':'Test de micrófono')+'</h2>'+
            '<p class="muted">Habla unos segundos para comprobar que la diadema graba y se escucha correctamente.</p>'+
          '</div>'+
          '<button id="closeAudioTest" class="audio-close" aria-label="Cerrar">×</button>'+
        '</div>'+
        '<div class="audio-test-layout '+(withGuide?'':'audio-test-layout-single')+'">'+'<div id="standaloneAudio"></div>'+guide+'</div>'+
      '</div>'+
    '</div>';

  $("#closeAudioTest").onclick=close;

  audioTest(
    $("#standaloneAudio"),
    ()=>{}
  );
}

async function requireLot(){
  if(await activeLot())return true;

  alert("Primero crea o selecciona un lote.");
  openLots();

  return false;
}

function openForm(d={}){
  draft={...d};

  root.innerHTML=
    '<div class="modalback">'+
      '<div class="modal">'+
        '<h2 id="formTitle"></h2>'+
        '<p class="muted">ID y ticket son independientes. Tipo de lectura: <span id="readType"></span></p>'+
        '<div class="formgrid">'+
          '<div class="field">'+
            "<label>ID</label>"+
            '<input id="id" placeholder="Vacío = generar J001, J002…">'+
          "</div>"+
          '<div class="field">'+
            "<label>Ticket</label>"+
            '<input id="ticket">'+
          "</div>"+
          '<div class="field full">'+
            '<label>Test de audio</label>'+
            '<div id="audio">'+
              '<button id="test">🧪 Hacer test</button>'+
              "<button id=\"skip\">Omitir</button>"+
            "</div>"+
          "</div>"+
          '<div class="field">'+
            "<label>Fabricante</label>"+
            '<select id="man"></select>'+
          "</div>"+
          '<div class="field">'+
            "<label>Diagnóstico</label>"+
            '<select id="diag">'+
              "<option>Reparable</option>"+
              "<option>No reparable</option>"+
            "</select>"+
          "</div>"+
          '<div class="field">'+
            "<label>Empaque</label>"+
            '<select id="pack">'+
              "<option>Listo para empacar</option>"+
              "<option>Empacado</option>"+
            "</select>"+
          "</div>"+
          '<div class="field">'+
            "<label>Fecha de recepción</label>"+
            '<input id="receivedDate" type="date">'+
          "</div>"+
          '<div class="field">'+
            "<label>Estado de entrega</label>"+
            '<select id="deliveryStatus"><option value="pending">Pendiente</option><option value="delivered">Entregada</option></select>'+
          "</div>"+
          '<div class="field">'+
            "<label>Fecha de entrega</label>"+
            '<input id="deliveryDate" type="date">'+
          "</div>"+
          '<div class="field">'+
            "<label>Garantía</label>"+
            '<select id="warrantyMonths"><option value="2">2 meses</option><option value="3">3 meses</option><option value="4">4 meses</option><option value="5">5 meses</option><option value="6">6 meses</option></select>'+
          "</div>"+
          '<div class="field">'+
            "<label>Garantía hasta</label>"+
            '<input id="warrantyUntil" type="date" readonly>'+
          "</div>"+
          '<div class="field full">'+
            "<label>Motivos de no reparación</label>"+
            '<div id="reasons" class="choices"></div>'+
          "</div>"+
          '<div class="field full">'+
            "<label>Reparaciones (máximo 3)</label>"+
            '<div id="repairs" class="choices"></div>'+
          "</div>"+
          '<div class="field full">'+
            "<label>Observaciones</label>"+
            '<textarea id="obs"></textarea>'+
          "</div>"+
        "</div>"+
        '<div class="modal-actions">'+
          '<button id="cancel">Cancelar</button>'+
          '<button id="save" class="primary">Guardar</button>'+
        "</div>"+
      "</div>"+
    "</div>";

  $("#formTitle").textContent=
    editing
      ?"Editar diadema"
      :"Registrar diadema";

  $("#readType").textContent=d.readType||"Manual";

  $("#id").value=d.unitId||"";
  $("#ticket").value=d.ticket||"";

  $("#man").innerHTML=CONFIG.manufacturers
    .map(x=>"<option>"+esc(x)+"</option>")
    .join("");

  $("#man").value=d.manufacturer||"Pendiente";

  $("#diag").value=
    d.diagnosis==="No reparable"
      ?"No reparable"
      :"Reparable";

  $("#pack").value=
    d.packaging||"Listo para empacar";

  $("#receivedDate").value=d.lotDate||"";
  $("#deliveryStatus").value=d.delivered===true?"delivered":"pending";
  $("#deliveryDate").value=d.deliveryDate||"";
  $("#warrantyMonths").value=String(d.warrantyMonths||2);
  const updateWarrantyDate=()=>{
    const delivered=$("#deliveryStatus").value==="delivered";
    const date=$("#deliveryDate").value;
    if(!delivered||!date){$("#warrantyUntil").value="";return;}
    const result=new Date(date+"T00:00:00");
    const day=result.getDate();
    result.setDate(1);
    result.setMonth(result.getMonth()+Number($("#warrantyMonths").value||2));
    const lastDay=new Date(result.getFullYear(),result.getMonth()+1,0).getDate();
    result.setDate(Math.min(day,lastDay));
    $("#warrantyUntil").value=[result.getFullYear(),String(result.getMonth()+1).padStart(2,"0"),String(result.getDate()).padStart(2,"0")].join("-");
  };
  $("#deliveryStatus").onchange=()=>{
    if($("#deliveryStatus").value==="delivered"&&!$("#deliveryDate").value){$("#deliveryDate").value=new Date().toISOString().slice(0,10);}
    if($("#deliveryStatus").value==="pending")$("#deliveryDate").value="";
    updateWarrantyDate();
  };
  $("#deliveryDate").onchange=updateWarrantyDate;
  $("#warrantyMonths").onchange=updateWarrantyDate;
  activeLot().then(lot=>{if(!$("#receivedDate")?.value&&lot)$("#receivedDate").value=lot.date||"";});
  updateWarrantyDate();

  $("#reasons").innerHTML=CONFIG.noRepairReasons
    .map(x=>
      '<label><input type="checkbox" name="reason" value="'+
      esc(x)+
      '"> '+
      esc(x)+
      "</label>")
    .join("");

  $("#repairs").innerHTML=CONFIG.repairOptions
    .map(x=>
      '<label><input type="checkbox" name="repair" value="'+
      esc(x)+
      '"> '+
      esc(x)+
      "</label>")
    .join("");

  (d.noRepairReasons||[]).forEach(v=>{
    const e=[...root.querySelectorAll('input[name="reason"]')]
      .find(i=>i.value===v);

    if(e)e.checked=true;
  });

  (d.repairs||[]).forEach(v=>{
    const e=[...root.querySelectorAll('input[name="repair"]')]
      .find(i=>i.value===v);

    if(e)e.checked=true;
  });

  const rules=()=>{
    const no=$("#diag").value==="No reparable";

    const m=[
      ...root.querySelectorAll('input[name="repair"]')
    ].find(i=>i.value==="Mantenimiento");

    if(no){
      $("#pack").value="Listo para empacar";

      root.querySelectorAll('input[name="repair"]')
        .forEach(i=>i.checked=false);

      if(m)m.checked=false;
    }else if(m){
      m.checked=true;
    }

    root.querySelectorAll('input[name="reason"]')
      .forEach(i=>i.disabled=!no);

    root.querySelectorAll('input[name="repair"]')
      .forEach(i=>{
        i.disabled=no||i.value==="Mantenimiento";
      });

    if(!no&&m)m.disabled=true;

    $("#reasons").classList.toggle("hidden",!no);
    $("#repairs").classList.toggle("hidden",no);
  };

  $("#diag").onchange=rules;

  root.querySelectorAll('input[name="repair"]')
    .forEach(i=>{
      i.onchange=()=>{
        if(
          root.querySelectorAll(
            'input[name="repair"]:checked'
          ).length>3
        ){
          i.checked=false;

          alert(
            "Máximo 3 trabajos de reparación, contando Mantenimiento."
          );
        }
      };
    });

  $("#cancel").onclick=close;

  $("#test").onclick=()=>
    audioTest(
      $("#audio"),
      v=>$("#audio").dataset.result=v
    );

  $("#skip").onclick=()=>
    $("#audio").dataset.result="Omitido";

  $("#save").onclick=async()=>{
    try{
      const id=$("#id").value.trim();
      const no=$("#diag").value==="No reparable";

      const rs=[
        ...root.querySelectorAll(
          'input[name="repair"]:checked'
        )
      ].map(i=>i.value);

      const nr=[
        ...root.querySelectorAll(
          'input[name="reason"]:checked'
        )
      ].map(i=>i.value);

      if(
        !no&&
        !rs.includes("Mantenimiento")
      ){
        return alert(
          "Mantenimiento es obligatorio."
        );
      }

      if(no&&rs.length){
        return alert(
          "Una unidad no reparable no puede tener trabajos de reparación."
        );
      }

      if(no&&!nr.length){
        return alert(
          "Selecciona al menos un motivo de no reparación."
        );
      }

      const all=await getAll(CONFIG.store);

      if(
        id&&
        all.some(
          x=>x.unitId===id&&
          x.id!==editing?.id
        )
      ){
        return alert(
          "Este número ya se encuentra registrado. Puedes continuar con la siguiente diadema."
        );
      }

      const saved=await saveUnit({
        unitId:id,
        ticket:$("#ticket").value.trim(),
        manufacturer:$("#man").value,
        diagnosis:$("#diag").value,
        repairable:!no,
        noRepairReasons:no?nr:[],
        repairs:no?[]:rs,
        observations:$("#obs").value.trim(),
        packaging:$("#pack").value,
        lotDate:$("#receivedDate").value||d.lotDate||"",
        delivered:$("#deliveryStatus").value==="delivered",
        deliveryStatus:$("#deliveryStatus").value==="delivered"?"Entregada":"Pendiente",
        deliveryDate:$("#deliveryStatus").value==="delivered"?$("#deliveryDate").value:"",
        warrantyMonths:Number($("#warrantyMonths").value||2),
        warrantyUntil:$("#warrantyUntil").value||"",
        readType:draft.readType||"Manual",
        audioTest:
          $("#audio").dataset.result||
          d.audioTest||
          "Omitido"
      },editing?.id);

      editing=null;

      close();

      await refresh();

      await sync.broadcastSnapshot(
        await getSnapshot()
      );

      alert(
        "Unidad guardada: "+saved.unitId
      );
    }catch(e){
      alert(
        e.message||
        "No se pudo guardar."
      );
    }
  };

  rules();
}

async function openLots(){
  const l=await listLots();

  const list=l.map(x=>
    '<div class="lot-item">'+
      "<div>"+
        "<b>"+esc(x.name)+"</b><br>"+
        '<span class="muted">'+
          x.date+
          " · "+
          esc(x.status||"active")+
        "</span>"+
      "</div>"+
      "<div>"+
        '<button data-l="'+x.id+'">Abrir</button> '+
        '<button data-e="'+x.id+'">Editar</button> '+
        '<button class="danger" data-d="'+x.id+'">Eliminar</button>'+
      "</div>"+
    "</div>"
  ).join("");

  root.innerHTML=
    '<div class="modalback">'+
      '<div class="modal">'+
        "<h2>Lotes</h2>"+
        '<div class="lot-list">'+
          list+
        "</div>"+
        '<div class="modal-actions">'+
          '<button id="new" class="primary">Nuevo lote</button>'+
          '<button id="x">Cerrar</button>'+
        "</div>"+
      "</div>"+
    "</div>";

  $("#x").onclick=close;

  $("#new").onclick=()=>lotForm();

  root.querySelectorAll("[data-l]")
    .forEach(b=>{
      b.onclick=async()=>{
        await selectLot(b.dataset.l);
        close();
        refresh();
      };
    });

  root.querySelectorAll("[data-e]")
    .forEach(b=>{
      b.onclick=async()=>{
        const x=await getByKey(
          CONFIG.lots,
          b.dataset.e
        );

        lotForm(x);
      };
    });

  root.querySelectorAll("[data-d]")
    .forEach(b=>{
      b.onclick=async()=>{
        if(
          confirm(
            "Eliminar lote y sus unidades? Esta acción no se puede deshacer."
          )
        ){
          await deleteLot(b.dataset.d);

          await sync.broadcastSnapshot(
            await getSnapshot()
          );

          openLots();
          refresh();
        }
      };
    });
}

function lotForm(l={}){
  root.innerHTML=
    '<div class="modalback">'+
      '<div class="modal">'+
        "<h2>"+
          (l.id?"Editar lote":"Nuevo lote")+
        "</h2>"+
        '<div class="formgrid">'+
          '<div class="field">'+
            "<label>Nombre</label>"+
            '<input id="ln" value="'+esc(l.name)+'">'+
          "</div>"+
          '<div class="field">'+
            "<label>Fecha del lote</label>"+
            '<input id="ld" type="date" value="'+
              esc(
                l.date||
                new Date()
                  .toISOString()
                  .slice(0,10)
              )+
            '">'+
          "</div>"+
          '<div class="field">'+
            "<label>Estado</label>"+
            '<select id="ls">'+
              '<option value="active">Activo</option>'+
              '<option value="closed">Cerrado</option>'+
            "</select>"+
          "</div>"+
          '<div class="field full">'+
            "<label>Observaciones</label>"+
            '<textarea id="lo">'+
              esc(l.observations)+
            "</textarea>"+
          "</div>"+
        "</div>"+
        '<div class="modal-actions">'+
          '<button id="lc">Cancelar</button>'+
          '<button id="lok" class="primary">Guardar lote</button>'+
        "</div>"+
      "</div>"+
    "</div>";

  $("#ls").value=l.status||"active";

  $("#lc").onclick=openLots;

  $("#lok").onclick=async()=>{
    const name=$("#ln").value.trim();

    if(!name){
      return alert("Escribe un nombre.");
    }

    if(l.id){
      await updateLot({
        ...l,
        name,
        date:$("#ld").value,
        status:$("#ls").value,
        observations:$("#lo").value.trim()
      });
    }else{
      await createLot(
        name,
        $("#ld").value,
        $("#lo").value.trim()
      );
    }

    await sync.broadcastSnapshot(
      await getSnapshot()
    );

    close();
    refresh();
  };
}

async function getSnapshot(){
  return{
    units:await getAll(CONFIG.store),
    lots:await getAll(CONFIG.lots),
    meta:await getAll(CONFIG.meta)
  };
}

async function applySnapshot(snapshot){
  if(
    !snapshot||
    !Array.isArray(snapshot.units)||
    !Array.isArray(snapshot.lots)||
    !Array.isArray(snapshot.meta)
  ){
    throw new Error(
      "Datos de sincronización inválidos."
    );
  }

  for(
    const store of [
      CONFIG.store,
      CONFIG.lots,
      CONFIG.meta
    ]
  ){
    const current=await getAll(store);
    await Promise.all(
      current.map(x=>
        del(
          store,
          x.id??x.key
        )
      )
    );
  }

  await Promise.all([
    ...snapshot.units.map(x=>put(CONFIG.store,x)),
    ...snapshot.lots.map(x=>put(CONFIG.lots,x)),
    ...snapshot.meta.map(x=>put(CONFIG.meta,x))
  ]);

  await refresh();
}

/* =========================================================
   SINCRONIZACIÓN
   Esta sección se conserva del A viejo.
   ========================================================= */

function syncStatusText(info){
  if(!info)return"Sin conexión";
  if(info.status==="ready")return info.detail||"Listo para conectar.";
  if(info.status==="checking")return info.detail||"Comparando los datos…";
  if(info.status==="synced")return info.detail||"Datos sincronizados.";
  if(info.status==="connected")return info.detail||"Dispositivo conectado.";
  if(info.status==="disconnected")return info.detail||"Sin conexión.";
  if(info.status==="error")return info.detail||"Error de sincronización.";
  return info.detail||info.status||"Sincronización";
}
function updateSyncVisual(info){
  const status=$("#syncStatus");
  if(!status)return;
  status.textContent=syncStatusText(info);
  status.classList.toggle("connected",info.status==="synced"||info.status==="connected");
}
async function handleSyncStatus(info){
  updateSyncVisual(info);
  if(info.status==="synced")await refresh();
}
function openSync(){
  root.innerHTML='<div class="modalback"><div class="modal"><h2>🔗 Sincronizar PC ↔ móvil</h2><div id="syncStatus" class="sync-status">Sin conexión</div><div class="field"><label>Código de este dispositivo</label><div id="syncCode" class="sync-code">Aún no generado</div><div id="syncQr" class="sync-qr"></div><small class="sync-help">Genera el código en el dispositivo que tenga el inventario. El otro puede escanearlo o introducirlo.</small></div><div class="modal-actions" style="position:static"><button id="makeSync" class="primary">Generar código</button><button id="scanSync">📷 Escanear QR</button><button id="disconnectSync" class="secondary">Desconectar</button><button id="closeSync">Cerrar</button></div><div class="field"><label>Código del otro dispositivo</label><input id="joinCode" placeholder="Pega aquí el código"></div><div class="modal-actions" style="position:static"><button id="joinSync" class="primary">Conectar</button></div><div class="sync-camera-wrap"><video id="syncCamera" class="sync-camera hidden" autoplay playsinline muted></video><button id="toggleFlash" class="flash-btn hidden" type="button">🔦 Activar flash</button></div></div></div>';
  updateSyncVisual({status:sync.isConnected()?"connected":"disconnected",detail:sync.isConnected()?"Dispositivo conectado.":"Sin conexión."});
  const code=$("#syncCode"),qr=$("#syncQr"),video=$("#syncCamera"),flash=$("#toggleFlash");
  let cam=null,flashOn=false;
  const stopCamera=()=>{cam?.getTracks().forEach(t=>t.stop());cam=null;flashOn=false;flash?.classList.add("hidden");if(flash)flash.textContent="🔦 Activar flash";video.classList.add("hidden");video.srcObject=null};
  const setupFlash=()=>{const track=cam?.getVideoTracks?.()[0],caps=track?.getCapabilities?.()||{};flash.classList.remove("hidden");if(caps.torch){flash.disabled=false;flash.textContent="🔦 Activar flash"}else{flash.disabled=true;flash.textContent="🔦 Flash no disponible"}};
  const setFlash=async on=>{const track=cam?.getVideoTracks?.()[0];if(!track)return;try{await track.applyConstraints({advanced:[{torch:on}]});flashOn=on;flash.textContent=on?"💡 Apagar flash":"🔦 Activar flash"}catch{flashOn=false;flash.textContent="🔦 Activar flash";alert("El flash no está disponible en esta cámara o navegador.")}};
  $("#makeSync").onclick=async()=>{try{const id=await sync.host();code.textContent=id;qr.innerHTML="";if(window.QRCode)new QRCode(qr,{text:id,width:190,height:190});updateSyncVisual({status:"ready",detail:"Código listo. Escanéalo desde el otro dispositivo."})}catch(e){updateSyncVisual({status:"error",detail:e.message||"No se pudo generar el código."})}};
  $("#joinSync").onclick=async()=>{try{const id=$("#joinCode").value.trim();if(!id)return updateSyncVisual({status:"error",detail:"Escribe o escanea el código del otro dispositivo."});await sync.join(id);updateSyncVisual({status:"ready",detail:"Conectando…"})}catch(e){updateSyncVisual({status:"error",detail:e.message||"No se pudo conectar."})}};
  $("#disconnectSync").onclick=()=>{sync.disconnect();updateSyncVisual({status:"disconnected",detail:"Sin conexión."})};
  $("#closeSync").onclick=()=>{stopCamera();close()};
  $("#toggleFlash").onclick=()=>setFlash(!flashOn);
  $("#scanSync").onclick=async()=>{try{if(!("BarcodeDetector"in window))return alert("Este navegador no permite escanear QR automáticamente. Usa la casilla de código.");stopCamera();cam=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"}}});video.classList.remove("hidden");video.srcObject=cam;await video.play();setupFlash();const detector=new BarcodeDetector({formats:["qr_code"]});let tries=0;const loop=async()=>{if(video.classList.contains("hidden"))return;try{const found=await detector.detect(video);if(found[0]?.rawValue){const value=found[0].rawValue;$("#joinCode").value=value;stopCamera();await sync.join(value);updateSyncVisual({status:"ready",detail:"Conectando…"});return}}catch{}if(++tries<300)setTimeout(loop,200);else stopCamera()};loop()}catch{alert("No se pudo abrir la cámara para escanear el código.")}};
}
$("#scanBtn").onclick=async()=>{
  if(!await requireLot())return;

  scanIdentification(r=>{
    if(r.type==="manual"){
      closeScanner();
      openForm({
        readType:"Manual"
      });
    }else{
      openForm({
        unitId:r.value,
        readType:r.type
      });
    }
  });
};

$("#audioTestBtn").onclick=()=>openAudioTest(false);
$("#audioGuideBtn").onclick=()=>openAudioTest(true);

$("#manualBtn").onclick=async()=>{
  if(await requireLot()){
    openForm({
      readType:"Manual"
    });
  }
};

$("#inventoryBody").onclick=async e=>{
  const a=e.target.closest("[data-edit]");
  const b=e.target.closest("[data-del]");

  if(a){
    editing=await getByKey(
      CONFIG.store,
      a.dataset.edit
    );

    openForm(editing);
  }

  if(
    b&&
    confirm(
      "¿Eliminar? El ID quedará reservado y no se reutilizará."
    )
  ){
    await removeUnit(
      b.dataset.del
    );

    await refresh();

    await sync.broadcastSnapshot(
      await getSnapshot()
    );
  }
};

$("#lotBtn").onclick=openLots;
$("#syncBtn").onclick=openSync;
$("#idFilter").oninput=refresh;

$("#copyBtn").onclick=async()=>{
  const u=await unitsForActiveLot();

  const txt=u.map(x=>[
    x.sequence,
    x.ticket,
    x.unitId,
    x.manufacturer,
    x.lotDate,
    x.readType,
    x.diagnosis,
    x.repairable===true
      ?"Sí"
      :x.repairable===false
        ?"No"
        :"",
    (x.repairs||[]).join(" + "),
    x.observations,
    x.packaging
  ].join("\t")).join("\n");

  try{
    await navigator.clipboard.writeText(txt);
    alert("Tabla copiada.");
  }catch{
    const t=document.createElement("textarea");

    t.value=txt;

    document.body.append(t);

    t.select();

    document.execCommand("copy");

    t.remove();

    alert("Tabla copiada.");
  }
};

$("#exportExcel").onclick=async()=>
  exportExcel(
    await unitsForActiveLot()
  );

$("#importExcel").onclick=async()=>{
  if(!(await requireLot())){
    return;
  }

  $("#importExcelFile").click();
};

$("#importExcelFile").onchange=async e=>{
  const input=e.target;
  const file=input.files?.[0];

  if(!file){
    return;
  }

  try{
    const lot=await activeLot();

    if(!lot){
      throw new Error("Primero crea o selecciona un lote.");
    }

    const result=await readExcelFile(file);
    const current=await unitsForActiveLot();
    const normalizeUnitId=value=>String(value??"").trim().toLocaleLowerCase();
    const currentIds=new Set(
      current.map(x=>normalizeUnitId(x.unitId)).filter(Boolean)
    );

    // Consultar todas las unidades, no solo las del lote activo: un ID que
    // ya existe en otro lote debe omitirse, no confundirse con un ID eliminado.
    const allUnits=await getAll(CONFIG.store);
    const existingIds=new Set(
      allUnits.map(x=>normalizeUnitId(x.unitId)).filter(Boolean)
    );

    const history=await getByKey(
      CONFIG.meta,
      "usedIds"
    );

    const usedIds=new Set(
      (Array.isArray(history?.value)?history.value:[])
        .map(normalizeUnitId)
        .filter(Boolean)
    );

    const pending=[];
    let skipped=0;

    for(let i=0;i<result.rows.length;i++){
      const imported=rowToUnit(result.rows[i],i,lot);
      const unitId=String(imported.unitId||"").trim();
      const normalizedId=normalizeUnitId(unitId);

      if(currentIds.has(normalizedId)||existingIds.has(normalizedId)){
        skipped++;
        continue;
      }

      if(usedIds.has(normalizedId)){
        throw new Error(
          `El ID "${unitId}" figura en el historial de IDs reservados (por ejemplo, una unidad eliminada). No se importó el archivo para evitar reutilizarlo. Si este Excel contiene una unidad que debe recuperarse, revisa una copia de seguridad o asigna un ID nuevo y único.`
        );
      }

      currentIds.add(normalizedId);
      existingIds.add(normalizedId);
      usedIds.add(normalizedId);
      pending.push(imported);
    }

    for(const imported of pending){
      await saveUnit({
        ...imported,
        sequence:undefined
      });
    }

    const added=pending.length;

    await refresh();

    await sync.broadcastSnapshot(
      await getSnapshot()
    );

    alert(
      "Excel importado correctamente.\n"+
      "Unidades añadidas: "+added+
      "\nUnidades ya existentes omitidas: "+skipped+
      "\nEl conteo continúa desde la última unidad."
    );
  }catch(e){
    alert(
      "No se pudo importar el Excel: "+
      (e.message||"archivo inválido.")
    );
  }finally{
    input.value="";
  }
};

$("#exportCsv").onclick=async()=>
  exportCsv(
    await unitsForActiveLot()
  );

$("#exportJson").onclick=async()=>{
  const lot=await activeLot();
  const u=await unitsForActiveLot();
  const used=await getByKey(
    CONFIG.meta,
    "usedIds"
  );
  const deleted=await getByKey(
    CONFIG.meta,
    "deletedIds"
  );
  const seq=await getByKey(
    CONFIG.meta,
    "lastSequence"
  );

  exportJson({
    version:CONFIG.appVersion,
    exportedAt:new Date().toISOString(),
    lot,
    units:u,
    history:{
      usedIds:used?.value||[],
      deletedIds:deleted?.value||[],
      lastSequence:seq?.value||0
    },
    configuration:CONFIG
  });
};

$("#importJson").onchange=async e=>{
  try{
    const f=e.target.files[0];

    if(!f)return;

    if(
      !confirm(
        "Importar reemplazará los datos actuales. ¿Continuar?"
      )
    ){
      return;
    }

    const p = JSON.parse(await f.text());
    if (!p || !Array.isArray(p.units) || !p.units.every(x => x && x.id && x.unitId) || (p.lot && !p.lot.id) || !p.history || !Array.isArray(p.history.usedIds) || !Array.isArray(p.history.deletedIds)) {
      throw new Error("El archivo no tiene una estructura de copia de seguridad válida. No se modificaron los datos actuales.");
    }
    const ids = p.units.map(x => String(x.unitId).trim());
    if (ids.some(id => !id) || new Set(ids).size !== ids.length) {
      throw new Error("La copia contiene IDs vacíos o duplicados. No se modificaron los datos actuales.");
    }

    for(
      const x of await getAll(CONFIG.store)
    ){
      await del(
        CONFIG.store,
        x.id
      );
    }

    for(
      const x of await getAll(CONFIG.lots)
    ){
      await del(
        CONFIG.lots,
        x.id
      );
    }

    if(p.lot){
      await put(
        CONFIG.lots,
        p.lot
      );

      await put(
        CONFIG.meta,
        {
          key:"activeLot",
          value:p.lot.id
        }
      );
    }else{
      await del(
        CONFIG.meta,
        "activeLot"
      );
    }

    for(const x of p.units||[]){
      await put(
        CONFIG.store,
        x
      );
    }

    if(p.history){
      await put(
        CONFIG.meta,
        {
          key:"usedIds",
          value:p.history.usedIds||[]
        }
      );

      await put(
        CONFIG.meta,
        {
          key:"deletedIds",
          value:p.history.deletedIds||[]
        }
      );

      await put(
        CONFIG.meta,
        {
          key:"lastSequence",
          value:p.history.lastSequence||0
        }
      );
    }

    location.reload();
  }catch(e){
    alert(
      "JSON inválido o incompatible: "+
      e.message
    );
  }

  e.target.value="";
};

/*
 * Configuración ÚNICA de sincronización.
 *
 * Se mantiene una sola configuración para que
 * el onStatus de la ventana no sea reemplazado.
 */
sync.configure({
  getSnapshot,
  applySnapshot,
  onStatus:handleSyncStatus
});

setInterval(()=>{
  if(
    document.querySelector(
      "#inventoryBody"
    )
  ){
    refresh();
  }
},15000);

refresh();

if("serviceWorker" in navigator){
  navigator.serviceWorker
    .register("./sw.js")
    .catch(()=>{});
}