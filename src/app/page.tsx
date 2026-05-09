export default function Home() {
  return (
    <main className="min-h-screen bg-neutral-900 flex items-center justify-center">
      <div className="text-center px-6">
        <h1 className="text-5xl font-bold text-white mb-4">Next.js Starter</h1>
        <p className="text-neutral-400 text-lg mb-8 max-w-md mx-auto">
          A minimal Next.js 16 template with TypeScript and Tailwind CSS 4. Ready to build anything.
        </p>
        <div className="flex gap-4 justify-center">
          <a
            href="https://nextjs.org/docs"
            target="_blank"
            rel="noopener noreferrer"
            className="px-5 py-2.5 bg-white text-neutral-900 font-medium rounded-lg hover:bg-neutral-200 transition-colors"
          >
            Documentation
          </a>
          <a
            href="https://github.com"
            target="_blank"
            rel="noopener noreferrer"
            className="px-5 py-2.5 border border-neutral-600 text-white font-medium rounded-lg hover:border-neutral-400 transition-colors"
          >
            GitHub
          </a>
        </div>
      </div>
    </main>
  );
}
