import Link from "next/link";

export default function HomePage() {
  return (
    <main className="container">
      <section className="card">
        <h1>Conversor de Remessa</h1>
        <p>Modulo implementado: Pagamento Bradesco &gt; PIX (CNAB240).</p>
        <Link href="/bradesco-transferencia" className="linkButton">
          Abrir modulo Bradesco PIX
        </Link>
      </section>
    </main>
  );
}
