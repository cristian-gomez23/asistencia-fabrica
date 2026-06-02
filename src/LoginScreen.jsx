import { useState, useEffect } from "react";

/* ── Google Fonts ── */
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href =
  "https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=DM+Sans:wght@300;400;500&display=swap";
if (!document.head.querySelector("link[data-login-font]")) {
  fontLink.dataset.loginFont = "1";
  document.head.appendChild(fontLink);
}

/* ──────────────────────────────────────────────────────────────
   Supabase — tomá las variables de tu .env
   VITE_SUPABASE_URL  y  VITE_SUPABASE_ANON_KEY
   ────────────────────────────────────────────────────────────── */
const SB_URL  = import.meta.env.VITE_SUPABASE_URL;
const SB_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY;

async function supabaseSignIn(email, password) {
  const res = await fetch(`${SB_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "apikey": SB_ANON,
    },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json();
  if (!res.ok) {
    // Supabase devuelve error_description en inglés; traducimos los más comunes
    const msg = data.error_description || data.msg || "Error al iniciar sesión.";
    if (msg.includes("Invalid login")) throw new Error("Correo o contraseña incorrectos.");
    if (msg.includes("Email not confirmed")) throw new Error("Confirmá tu correo antes de ingresar.");
    throw new Error(msg);
  }
  return data; // { access_token, user, ... }
}

/* ──────────────────────────────────────────────────────────────
   Favicon con imagen real
   ────────────────────────────────────────────────────────────── */
function setFavicon() {
  document.title = "PYG — Sistema RRHH";
  // Buscar o crear el tag <link rel="icon">
  let link = document.querySelector("link[rel~='icon']");
  if (!link) {
    link = document.createElement("link");
    link.rel = "icon";
    document.head.appendChild(link);
  }
  link.type = "image/png";
  link.href = "/PyG-logo_mini.png";
}

/* ──────────────────────────────────────────────────────────────
   Componente principal
   ────────────────────────────────────────────────────────────── */
export default function LoginScreen({ onLogin }) {
  const [email,    setEmail]    = useState("");
  const [clave,    setClave]    = useState("");
  const [error,    setError]    = useState("");
  const [loading,  setLoading]  = useState(false);
  const [showPass, setShowPass] = useState(false);
  const [mounted,  setMounted]  = useState(false);

  useEffect(() => {
    setFavicon();
    setTimeout(() => setMounted(true), 60);
  }, []);

  const handleLogin = async (e) => {
    e?.preventDefault();
    if (!email.trim() || !clave.trim()) {
      setError("Completá el correo y la contraseña.");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const session = await supabaseSignIn(email.trim(), clave);
      onLogin({
        email:    session.user.email,
        userId:   session.user.id,
        token:    session.access_token,
        // Supabase permite guardar metadata extra en user_metadata
        rol:      session.user.user_metadata?.rol || "Usuario",
        nombre:   session.user.user_metadata?.nombre || session.user.email,
      });
    } catch (err) {
      setError(err.message);
      setLoading(false);
    }
  };

  const S = {
    root: {
      position: "fixed", inset: 0,
      background: "linear-gradient(135deg, #0d2145 0%, #1a3a6b 45%, #2756a0 100%)",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'DM Sans', system-ui, sans-serif",
      overflow: "hidden",
    },
    bgLine: {
      position: "absolute", inset: 0,
      backgroundImage: "repeating-linear-gradient(0deg,transparent,transparent 59px,rgba(255,255,255,0.025) 60px),repeating-linear-gradient(90deg,transparent,transparent 59px,rgba(255,255,255,0.025) 60px)",
      pointerEvents: "none",
    },
    bgCircle1: {
      position: "absolute", width: 520, height: 520, borderRadius: "50%",
      background: "radial-gradient(circle,rgba(255,255,255,0.05) 0%,transparent 70%)",
      top: -140, right: -120, pointerEvents: "none",
    },
    bgCircle2: {
      position: "absolute", width: 360, height: 360, borderRadius: "50%",
      background: "radial-gradient(circle,rgba(61,107,158,0.35) 0%,transparent 70%)",
      bottom: -80, left: -80, pointerEvents: "none",
    },
    card: {
      position: "relative", background: "#ffffff", borderRadius: 20,
      boxShadow: "0 32px 80px rgba(0,0,0,0.35),0 0 0 1px rgba(255,255,255,0.08)",
      width: "100%", maxWidth: 420, padding: "44px 40px 40px",
      opacity: mounted ? 1 : 0,
      transform: mounted ? "translateY(0)" : "translateY(24px)",
      transition: "opacity 0.5s ease,transform 0.5s ease",
    },
    header: {
      display: "flex", flexDirection: "column", alignItems: "center", marginBottom: 32,
    },
    logoImg: {
      height: 72, width: "auto", objectFit: "contain", marginBottom: 16,
    },
    divider: {
      width: 48, height: 3,
      background: "linear-gradient(90deg,#1a3a6b,#3d6b9e)",
      borderRadius: 2, margin: "0 auto 12px",
    },
    subtitle: {
      fontSize: 13, color: "#7a8fa6", fontWeight: 400, textAlign: "center",
    },
    form:       { display: "flex", flexDirection: "column", gap: 18 },
    fieldGroup: { display: "flex", flexDirection: "column", gap: 7 },
    label: {
      fontSize: 11.5, fontWeight: 600, color: "#4a5a6a",
      letterSpacing: "0.06em", textTransform: "uppercase",
    },
    inputWrap:  { position: "relative", display: "flex", alignItems: "center" },
    input: {
      width: "100%", padding: "12px 14px",
      border: "1.5px solid #dde3ec", borderRadius: 10,
      fontSize: 14, color: "#1e2a38", background: "#f8fafc",
      fontFamily: "'DM Sans', system-ui, sans-serif",
      outline: "none", transition: "border-color 0.2s,box-shadow 0.2s",
      boxSizing: "border-box",
    },
    inputErr:  { borderColor: "#e53e3e", background: "#fff5f5" },
    eyeBtn: {
      position: "absolute", right: 12, background: "none", border: "none",
      cursor: "pointer", padding: 4, display: "flex", alignItems: "center",
      color: "#96a3b0",
    },
    errorBox: {
      display: "flex", alignItems: "center", gap: 8,
      background: "#fff5f5", border: "1px solid #feb2b2",
      borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#c53030",
    },
    submitBtn: {
      padding: "13px 0",
      background: loading ? "#5a7fa8" : "linear-gradient(135deg,#1a3a6b 0%,#2756a0 100%)",
      color: "#fff", border: "none", borderRadius: 10,
      fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 14,
      letterSpacing: "0.05em", cursor: loading ? "not-allowed" : "pointer",
      display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
      boxShadow: "0 4px 14px rgba(26,58,107,0.35)", transition: "opacity 0.2s",
    },
    footer: {
      marginTop: 28, paddingTop: 22, borderTop: "1px solid #edf0f5",
      display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
    },
    footerLogo: { height: 22, width: "auto", opacity: 0.5 },
    footerText: { fontSize: 11, color: "#b0bbc8" },
  };

  const focusIn  = e => { e.target.style.borderColor="#3d6b9e"; e.target.style.boxShadow="0 0 0 3px rgba(61,107,158,0.15)"; };
  const focusOut = e => { e.target.style.borderColor=error?"#e53e3e":"#dde3ec"; e.target.style.boxShadow="none"; };

  return (
    <div style={S.root}>
      <div style={S.bgLine} />
      <div style={S.bgCircle1} />
      <div style={S.bgCircle2} />

      <div style={S.card}>
        {/* Header con logo real */}
        <div style={S.header}>
          <img src="/PyG-logo.png" alt="PYG S.R.L." style={S.logoImg} />
          <div style={S.divider} />
          <div style={S.subtitle}>Recursos Humanos · Acceso al sistema</div>
        </div>

        {/* Formulario */}
        <form style={S.form} onSubmit={handleLogin}>

          <div style={S.fieldGroup}>
            <label style={S.label}>Correo electrónico</label>
            <div style={S.inputWrap}>
              <input
                style={{ ...S.input, ...(error ? S.inputErr : {}) }}
                type="email"
                autoComplete="email"
                placeholder="nombre@empresa.com"
                value={email}
                onChange={e => { setEmail(e.target.value); setError(""); }}
                disabled={loading}
                onFocus={focusIn}
                onBlur={focusOut}
              />
            </div>
          </div>

          <div style={S.fieldGroup}>
            <label style={S.label}>Contraseña</label>
            <div style={S.inputWrap}>
              <input
                style={{ ...S.input, paddingRight: 44, ...(error ? S.inputErr : {}) }}
                type={showPass ? "text" : "password"}
                autoComplete="current-password"
                placeholder="••••••••"
                value={clave}
                onChange={e => { setClave(e.target.value); setError(""); }}
                disabled={loading}
                onFocus={focusIn}
                onBlur={focusOut}
              />
              <button type="button" style={S.eyeBtn} onClick={() => setShowPass(v => !v)} tabIndex={-1}>
                {showPass ? (
                  <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                    <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
                    <line x1="1" y1="1" x2="23" y2="23"/>
                  </svg>
                ) : (
                  <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                    <circle cx="12" cy="12" r="3"/>
                  </svg>
                )}
              </button>
            </div>
          </div>

          {error && (
            <div style={S.errorBox}>
              <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <circle cx="12" cy="12" r="10"/>
                <line x1="12" y1="8" x2="12" y2="12"/>
                <line x1="12" y1="16" x2="12.01" y2="16"/>
              </svg>
              {error}
            </div>
          )}

          <button type="submit" style={S.submitBtn} disabled={loading}>
            {loading ? (
              <>
                <svg width="16" height="16" fill="none" stroke="white" strokeWidth="2.5" viewBox="0 0 24 24"
                  style={{ animation: "pyg-spin 0.8s linear infinite" }}>
                  <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/>
                </svg>
                Verificando…
              </>
            ) : (
              <>
                <svg width="16" height="16" fill="none" stroke="white" strokeWidth="2.5" viewBox="0 0 24 24">
                  <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
                  <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
                </svg>
                Ingresar al sistema
              </>
            )}
          </button>
        </form>

        {/* Footer */}
        <div style={S.footer}>
          <img src="/PyG-logo_mini.png" alt="PYG" style={S.footerLogo} />
          <span style={S.footerText}>
            PYG S.R.L. · RRHH &amp; Liquidaciones &copy; {new Date().getFullYear()}
          </span>
        </div>
      </div>

      <style>{`@keyframes pyg-spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
