/* =====================================================================
   Ponte Firebase para as páginas de mapa de obra
   ---------------------------------------------------------------------
   As páginas de mapa foram feitas para rodar como artifact da Claude e
   usam `claude.use("db" | "user" | "downloads")`. Este arquivo cria um
   objeto `window.claude` com a mesma interface, mas gravando no Firebase
   do CRM (Firestore: obras/<obra>/el/<id> e obras/<obra>/cfg/<doc>).

   Uso na página da obra, ANTES do script principal:
     <script>window.OBRA_ID = "ita-vilela";</script>
     <script src="../_ponte-firebase.js"></script>
   ===================================================================== */
(function () {
  const OBRA = window.OBRA_ID;
  const V = "10.12.0";
  const CONFIG = {
    apiKey: "AIzaSyDbNf_gcYCIbXevKrymSKVHWocEuRgdQLE",
    authDomain: "vieira-silva-crm.firebaseapp.com",
    projectId: "vieira-silva-crm",
    storageBucket: "vieira-silva-crm.firebasestorage.app",
    messagingSenderId: "571225974312",
    appId: "1:571225974312:web:4124ff2753b953fd60f71d"
  };

  // Carrega o Firebase, espera o login (o mesmo do CRM) e lê o perfil do usuário.
  const pronto = (async () => {
    if (!OBRA) throw { code: "sem-obra" };
    const [A, U, F] = await Promise.all([
      import(`https://www.gstatic.com/firebasejs/${V}/firebase-app.js`),
      import(`https://www.gstatic.com/firebasejs/${V}/firebase-auth.js`),
      import(`https://www.gstatic.com/firebasejs/${V}/firebase-firestore.js`)
    ]);
    const app = A.getApps().length ? A.getApps()[0] : A.initializeApp(CONFIG);
    const auth = U.getAuth(app);
    const db = F.getFirestore(app);
    const user = await new Promise(res => { const off = U.onAuthStateChanged(auth, u => { off(); res(u); }); });
    if (!user) throw { code: "unauthenticated" };
    let role = "operador";
    try { const d = await F.getDoc(F.doc(db, "usuarios", user.uid)); if (d.exists()) role = d.data().role || "operador"; } catch (e) {}
    return { F, db, user, role };
  })();

  pronto.catch(e => {
    if (e && e.code === "unauthenticated") avisoLogin();
  });

  function avisoLogin() {
    const show = () => {
      const d = document.createElement("div");
      d.style.cssText = "position:fixed;inset:0;background:rgba(7,26,74,.92);color:#fff;display:flex;align-items:center;justify-content:center;z-index:999;font-family:system-ui,sans-serif;text-align:center;padding:24px";
      d.innerHTML = '<div><div style="font-size:18px;font-weight:700;margin-bottom:8px">Entre no sistema para ver o andamento</div>' +
        '<a href="/" target="_top" style="color:#fff;background:#0E6FCC;padding:10px 18px;border-radius:8px;text-decoration:none;display:inline-block;margin-top:8px">Ir para o login</a></div>';
      document.body.appendChild(d);
    };
    if (document.body) show(); else addEventListener("DOMContentLoaded", show);
  }

  const ref = (F, db, path) => F.doc(db, "obras", OBRA, ...path.split("/"));
  const limpo = o => JSON.parse(JSON.stringify(o)); // Firestore não aceita undefined

  const dbApi = {
    collection(nome) {
      return {
        onSnapshot(cb, err) {
          let off = null, parado = false;
          pronto.then(({ F, db }) => {
            if (parado) return;
            off = F.onSnapshot(F.collection(db, "obras", OBRA, nome), snap => cb({
              metadata: { fromCache: snap.metadata.fromCache },
              size: snap.size,
              docs: snap.docs.map(d => ({ id: d.id, exists: true, data: () => d.data() }))
            }), e => err && err(e));
          }).catch(e => err && err(e));
          return () => { parado = true; off && off(); };
        }
      };
    },
    doc(path) {
      return {
        async set(data) {
          const { F, db } = await pronto;
          try { await F.setDoc(ref(F, db, path), limpo(data)); }
          catch (e) { throw (e && e.code === "permission-denied") ? { code: "invalid_argument" } : e; }
        },
        onSnapshot(cb, err) {
          let off = null, parado = false;
          pronto.then(({ F, db }) => {
            if (parado) return;
            off = F.onSnapshot(ref(F, db, path), d => cb({ exists: d.exists(), data: () => d.data() }), e => err && err(e));
          }).catch(e => err && err(e));
          return () => { parado = true; off && off(); };
        }
      };
    }
  };

  const nomes = {};
  const userApi = {
    async id() { return (await pronto).user.uid; },
    // Visualizador só lê; admin e operador editam.
    can(perm) { return perm === "data.write" ? (window.__obraRole ? window.__obraRole !== "visualizador" : null) : null; },
    async profiles(ids) {
      const { F, db } = await pronto, out = {};
      await Promise.all(ids.map(async id => {
        if (!(id in nomes)) {
          try { const d = await F.getDoc(F.doc(db, "usuarios", id)); nomes[id] = d.exists() ? (d.data().nome || d.data().email || "") : ""; }
          catch (e) { nomes[id] = ""; }
        }
        out[id] = { name: nomes[id] };
      }));
      return out;
    }
  };

  const downloadsApi = {
    async save({ filename, data }) {
      const blob = new Blob([data], { type: "text/csv;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    }
  };

  window.claude = {
    async use(nome) {
      const p = await pronto;            // lança erro se não estiver logado
      window.__obraRole = p.role;
      if (nome === "db") return dbApi;
      if (nome === "user") return userApi;
      if (nome === "downloads") return downloadsApi;
      return null;
    }
  };
})();
