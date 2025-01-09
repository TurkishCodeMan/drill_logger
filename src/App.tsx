import { ExcelForm } from './components/ExcelForm'
import { Toaster } from 'sonner'

function App() {
  return (
    <div className="min-h-screen bg-background">
      <main className="container mx-auto py-10">
        <ExcelForm />
      </main>
      <Toaster />
    </div>
  )
}

export default App 