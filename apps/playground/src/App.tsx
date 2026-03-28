export function App() {
  return (
    <main
      style={{
        alignItems: "center",
        background: "#f3f4f6",
        color: "#101828",
        display: "grid",
        fontFamily: '"Aptos", "Segoe UI", "IBM Plex Sans", sans-serif',
        minHeight: "100vh",
        padding: "32px",
      }}
    >
      <section
        style={{
          background: "#ffffff",
          border: "1px solid #d0d5dd",
          borderRadius: "16px",
          boxShadow: "0 14px 40px rgba(15, 23, 42, 0.08)",
          maxWidth: "620px",
          padding: "28px",
        }}
      >
        <p
          style={{
            color: "#667085",
            fontSize: "12px",
            fontWeight: 700,
            letterSpacing: "0.08em",
            margin: 0,
            textTransform: "uppercase",
          }}
        >
          Deprecated surface
        </p>
        <h1 style={{ fontSize: "28px", lineHeight: 1.2, margin: "12px 0 8px" }}>
          bilig playground has been retired
        </h1>
        <p style={{ fontSize: "16px", lineHeight: 1.6, margin: 0 }}>
          Use the production web app in <code>@bilig/web</code>. This package is preserved only for
          historical reference and should not be used for active development or smoke tests.
        </p>
      </section>
    </main>
  );
}
