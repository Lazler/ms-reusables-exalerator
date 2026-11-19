(function(){
  const settings = {};
  let nextId = 100;
  function mkShape(o){
    const tags = {};
    const sh = {
      id: String(o.id ?? nextId++), name: o.name, type: o.type || "GeometricShape",
      width: o.width ?? 100, height: o.height ?? 50, left: 0, top: 0,
      load(){},
      fill: { type: "Solid", foregroundColor: o.fill || "#FFFFFF", transparency: 0, load(){},
        setSolidColor(c){ this.type="Solid"; this.foregroundColor=c; }, clear(){ this.type="NoFill"; } },
      lineFormat: { visible: true, color: "#000000", weight: 1, dashStyle: "Solid", load(){} },
      textFrame: { verticalAlignment: "Top", leftMargin: 7, rightMargin: 7, topMargin: 3, bottomMargin: 3, wordWrap: true, autoSizeSetting: "AutoSizeNone", load(){},
        textRange: { text: o.text || "", load(){}, font: { name:"Aptos", size:18, bold:false, italic:false, underline:"None", color:"#000000", load(){} },
          paragraphFormat: { horizontalAlignment: "Left", load(){} } } },
      tags: {
        add(k,v){ tags[k.toUpperCase()] = String(v); },
        delete(k){ if(!(k.toUpperCase() in tags)) throw new Error("ItemNotFound "+k); delete tags[k.toUpperCase()]; },
        getItemOrNullObject(k){ const K=k.toUpperCase(); return { load(){}, get isNullObject(){ return !(K in tags); }, get value(){ return tags[K]; } }; },
        _raw: tags,
      },
    };
    return sh;
  }
  const slides = [
    { id: "256#", shapes: [mkShape({id:2,name:"Title",type:"TextBox",text:"Offer for {{Client}}"}), mkShape({id:3,name:"Box",fill:"#FF0000",text:"Go"})] },
    { id: "257#", shapes: [mkShape({id:2,name:"Text",type:"TextBox",text:"{{Client}} – {{Project}}"}), mkShape({id:4,name:"Picture",type:"Image"})] },
  ];
  // A picture has no text frame -> accessing it throws on sync
  slides[1].shapes[1].textFrame = { get textRange(){ window.__badAccess = true; return { load(){ window.__pendingError = true; } }; } };
  let selection = { slide: 0, shapes: [] };
  const coll = (arr) => ({ items: arr, load(){}, getCount(){ return { value: arr.length }; }, getItemAt(i){ return arr[i]; } });
  const ctx = {
    presentation: {
      slides: { get items(){ return slides.map(s => ({ id: s.id, shapes: coll(s.shapes), setSelectedShapes(ids){ selection = { slide: slides.indexOf(s), shapes: ids }; } })); }, load(){},
        getItemOrNullObject(id){ const s = slides.find(x=>x.id===id); if (!s) return { isNullObject: true, load(){}, shapes: { load(){}, items: [] } }; return { isNullObject: false, id: s.id, load(){}, shapes: coll(s.shapes) }; },
        getItem(id){ const s = slides.find(x=>x.id===id); return { setSelectedShapes(ids){ selection = { slide: slides.indexOf(s), shapes: ids }; },
          shapes: { getItem(sid){ const sh = s.shapes.find(x => x.id === String(sid)); if (!sh) throw new Error("ItemNotFound"); return sh; } } }; } },
      getSelectedShapes(){ if (window.__selFail > 0) { window.__selFail--; window.__pendingError = true; } const s = slides[selection.slide]; return coll(s.shapes.filter(x => selection.shapes.includes(x.id))); },
      getSelectedSlides(){ return coll([{ id: slides[selection.slide].id }]); },
      getSelectedTextRangeOrNullObject(){ const s = slides[selection.slide].shapes.find(x => selection.shapes.includes(x.id)); if(!s) return { isNullObject:true }; const tr = s.textFrame.textRange; return { isNullObject:false, set text(v){ tr.text += v; } }; },
      setSelectedSlides(ids){ selection = { slide: slides.findIndex(s=>s.id===ids[0]), shapes: [] }; },
    },
    async sync(){ window.__syncs = (window.__syncs || 0) + 1; if (window.__pendingError){ window.__pendingError=false; throw Object.assign(new Error("GeneralException"),{code:"GeneralException",debugInfo:{errorLocation:"ShapeScopedCollection.getItem"}}); } },
  };
  const handlers = [];
  window.__pp = { slides, mkShape, select(slide, ids){ selection = { slide, shapes: ids.map(String) }; handlers.forEach(h=>h()); } };
  window.PowerPoint = { run: (fn) => fn(ctx) };
  window.Office = {
    HostType: { PowerPoint: "PowerPoint" }, AsyncResultStatus: { Succeeded: "succeeded" }, EventType: { DocumentSelectionChanged: "sel" },
    context: { requirements: { isSetSupported: () => true }, document: {
      settings: { get: k => settings[k], set: (k,v) => settings[k]=v, saveAsync: cb => cb({ status: "succeeded" }) },
      addHandlerAsync: (_, h) => handlers.push(h) } },
    onReady(cb){ setTimeout(() => cb({ host: "PowerPoint" }), 0); },
  };
})();
