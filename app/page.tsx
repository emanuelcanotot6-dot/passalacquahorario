'use client'

import { useEffect, useMemo, useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { toPng } from 'html-to-image'
import { Button } from '@/components/ui/button'
import { CalendarDays, Check, ChevronDown, Cloud, Download, ListFilter, Pencil, Printer, Search, ShieldAlert, ShieldCheck, UserCheck, UserX, Users, X, Zap, Clock, Trash2, Plus, Image as ImageIcon, FileSpreadsheet, Coffee, ArrowDownUp } from 'lucide-react'

type View = 'horario' | 'docentes' | 'ausencias' | 'preceptores' | 'paro' | 'suplencias'
type Status = 'Titular' | 'Provisional' | 'Titular Interino'
type Attendance = 'normal' | 'paro' | 'licencia' | 'no_contesta'
type CoberturaStatus = 'Cubierto por Decreto' | 'Sin cubrir'

type Block = { 
  id: string; 
  curso: string; 
  turno: 'Teoría' | 'Taller'; 
  tipo: 'Teoría' | 'Taller' | 'Recreo'; 
  dia: string; 
  horario: string; 
  materia: string; 
  docente: string; 
  situacionRevista: Status; 
  suplente?: string; 
  grupo?: string 
}

type AusenciaRecord = {
  id: string;
  curso: string;
  mes: string;
  fecha: string;
  materia: string;
  docenteAusente: string;
  estadoCobertura: CoberturaStatus;
  docenteDecreto: string;
}

type PreceptorAssignment = {
  curso: string;
  preceptorTitular: string;
  preceptorSuplente: string;
}

type ParoRecord = {
  id: string;
  fecha: string;
  curso: string;
  horario: string;
  materia: string;
  docente: string;
  estadoParo: Attendance;
  observaciones?: string;
  tipoPersonal: 'docente' | 'preceptor';
  grupo?: string;
}

type SuplenciaRecord = {
  id: string;
  fecha: string;
  curso: string;
  materia: string;
  titular: string;
  suplente: string;
  accion: 'Alta' | 'Baja';
}

const logo = '/logo.png'
const spreadsheetApi = 'https://script.google.com/macros/s/AKfycbz1LCenXovvxTnjkkYZeDsIbQe4uPmYtEEG5-K4lOfPhXZmO9ETm6IIKv70LdvrLhmm/exec'
const days = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
const courses = ['1°1°', '1°2°', '1°3°', '1°4°', '1°5°', '1°6°', '2°1°', '2°2°', '2°3°', '2°4°', '2°5°', '2°6°', '3°1°', '3°2°', '3°3°', '3°4°', '3°5°', '3°6°', '4°1°', '4°2°', '4°3°', '4°4°', '4°5°', '5°1°', '5°2°', '5°3°', '5°6°', '5°7°', '6°1°', '6°2°', '6°3°', '6°6°', '6°7°', '7°1°', '7°2°', '7°3°', '7°6°', '7°7°']

const seed: Block[] = []
const seedAusencias: AusenciaRecord[] = []
const seedSuplencias: SuplenciaRecord[] = []
const seedPreceptores: PreceptorAssignment[] = courses.map((c) => ({
  curso: c, preceptorTitular: '', preceptorSuplente: ''
}))

const statusStyle: Record<Status, string> = { Titular: 'text-[#0052cc]', Provisional: 'text-[#cc0000]', 'Titular Interino': 'text-slate-900' }

export default function Page() {
  const [view, setView] = useState<View>('horario')
  const [blocks, setBlocks] = useState<Block[]>(seed)
  const [ausencias, setAusencias] = useState<AusenciaRecord[]>(seedAusencias)
  const [historialSuplencias, setHistorialSuplencias] = useState<SuplenciaRecord[]>(seedSuplencias)
  const [preceptores, setPreceptores] = useState<PreceptorAssignment[]>(seedPreceptores)
  const [paroHistorial, setParoHistorial] = useState<ParoRecord[]>([])
  
  const [course, setCourse] = useState('1°1°')
  const [filter, setFilter] = useState<'Todos' | 'Teoría' | 'Taller'>('Todos')
  const [selectedDays, setSelectedDays] = useState(days)
  const [query, setQuery] = useState('')
  const [editing, setEditing] = useState<Block | null>(null)
  const [showRecreosModal, setShowRecreosModal] = useState(false)
  const [dirty, setDirty] = useState(0)
  const [saved, setSaved] = useState(false)
  const [syncing, setSyncing] = useState(false)
  const [syncMessage, setSyncMessage] = useState('Conectando con Google Drive…')

  useEffect(() => { syncSpreadsheet() }, [])

  const courseTimes = useMemo(() => {
    const courseBlocks = blocks.filter((b: Block) => b.curso === course)
    const timesSet = new Set(courseBlocks.map((b: Block) => b.horario))
    return Array.from(timesSet).sort((a: string, b: string) => a.localeCompare(b))
  }, [blocks, course])

  const visible = blocks.filter((b: Block) => b.curso === course && selectedDays.includes(b.dia) && (filter === 'Todos' || b.tipo === filter || b.tipo === 'Recreo'))
  
  const teachers = useMemo(() => {
    const allTeachers = [
      ...blocks.flatMap((b: Block) => [b.docente, b.suplente]),
      ...ausencias.flatMap((a: AusenciaRecord) => [a.docenteAusente, a.docenteDecreto])
    ].filter(Boolean).map(t => t!.trim());
    return [...new Set(allTeachers)].sort((a, b) => a.localeCompare(b));
  }, [blocks, ausencias])

  const searchResults = blocks.filter((b: Block) => (b.docente + ' ' + (b.suplente ?? '')).toLowerCase().includes(query.toLowerCase()))
  const dayBlocks = blocks.filter((b: Block) => b.tipo !== 'Recreo')

  function updateBlock(next: Block) { 
    setBlocks((current) => {
      const old = current.find(b => b.id === next.id);
      
      // Si el bloque existe, verificamos si cambió el suplente
      if (old) {
        const oldSup = (old.suplente || '').trim().toUpperCase();
        const newSup = (next.suplente || '').trim().toUpperCase();
        const today = new Date().toISOString().split('T')[0];

        if (oldSup !== newSup) {
          const newRecords: SuplenciaRecord[] = [];
          if (oldSup) {
            newRecords.push({ id: `HS-${Date.now()}-B`, fecha: today, curso: next.curso, materia: next.materia, titular: next.docente, suplente: oldSup, accion: 'Baja' });
          }
          if (newSup) {
            newRecords.push({ id: `HS-${Date.now()}-A`, fecha: today, curso: next.curso, materia: next.materia, titular: next.docente, suplente: newSup, accion: 'Alta' });
          }
          if (newRecords.length > 0) {
            setHistorialSuplencias(prev => [...newRecords, ...prev]);
          }
        }
        return current.map((b: Block) => b.id === next.id ? next : b);
      } else {
        // Es un bloque nuevo (Agregado desde el botón + Agregar)
        return [...current, next];
      }
    });

    setEditing(null);
    setDirty((n) => n + 1); 
  }

  function descargarBaseCompletaExcel() {
    const wb = XLSX.utils.book_new()
    const dataRows = blocks.map((b: Block) => ({
      Curso: b.curso, Turno: b.turno, Tipo: b.tipo, Día: b.dia, Grupo: b.grupo || '',
      Horario: b.horario, Materia: b.materia, Docente: b.docente, 'Situación Revista': b.situacionRevista, Suplente: b.suplente || ''
    }))
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataRows), 'DATA_CON_SITUACION')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(preceptores), 'PRECEPTORES')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paroHistorial), 'RELEVAMIENTO_PARO')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ausencias), 'AUSENCIAS_HISTORICO')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historialSuplencias), 'HISTORIAL_SUPLENCIAS')
    XLSX.writeFile(wb, 'HORARIOS-app-actualizado.xlsx')
  }
  
  async function syncSpreadsheet() {
    setSyncing(true)
    try {
      const response = await fetch(spreadsheetApi, { cache: 'no-store' })
      if (!response.ok) throw new Error('No se pudo leer la planilla')
      const payload = await response.json()
      const rows = Array.isArray(payload) ? payload : payload.data ?? payload.rows ?? payload.blocks ?? payload.values
      
      const read = (row: any, ...keys: string[]) => { const found = keys.find((key) => row?.[key] !== undefined && row?.[key] !== ''); return found ? row[found] : '' }
      
      if (Array.isArray(rows) && rows.length) {
        const normalized = rows.map((row: any, index: number) => {
          const doc = String(read(row, 'docente', 'Docente') || '').trim()
          let sit = String(read(row, 'situacionRevista', 'Situación Revista', 'Situacion Revista') || '').trim()
          if (!sit) sit = doc.toUpperCase().includes('TI 2026') ? 'Titular Interino' : 'Titular'
          return { 
            id: String(read(row, 'id', 'ID') || index + 1), 
            curso: String(read(row, 'curso', 'Curso') || (index === 0 ? '1°1°' : '')), 
            turno: (read(row, 'turno', 'Turno') || 'Teoría') as 'Teoría' | 'Taller', 
            tipo: (read(row, 'tipo', 'Tipo') || 'Teoría') as 'Teoría' | 'Taller' | 'Recreo', 
            dia: String(read(row, 'dia', 'Día', 'Dia') || ''), 
            horario: String(read(row, 'horario', 'Horario') || ''), 
            materia: String(read(row, 'materia', 'Materia') || ''), 
            docente: doc, situacionRevista: sit as Status, 
            suplente: String(read(row, 'suplente', 'Suplente') || '').trim(), 
            grupo: String(read(row, 'grupo', 'Grupo') || '') 
          }
        })

        // AUTO-LIMPIEZA DE DUPLICADOS: Si detecta múltiples bloques idénticos (mismo curso, dia, horario, materia y grupo) guarda solo 1
        const uniqueBlocksMap = new Map();
        normalized.forEach(b => {
          const key = `${b.curso}|${b.dia}|${b.horario}|${b.materia}|${b.grupo}`;
          if (!uniqueBlocksMap.has(key)) {
            uniqueBlocksMap.set(key, b);
          }
        });
        setBlocks(Array.from(uniqueBlocksMap.values()))
      }

      if (payload.ausencias && Array.isArray(payload.ausencias)) {
        const normAus = payload.ausencias.map((a: any, index: number) => ({
          id: String(read(a, 'id', 'ID') || `AUS-${index + 1}`),
          curso: String(read(a, 'curso', 'Curso') || '1°1°'),
          mes: String(read(a, 'mes', 'Mes') || 'Octubre'),
          fecha: String(read(a, 'fecha', 'Fecha', 'Fecha / Horario') || new Date().toISOString().split('T')[0]),
          materia: String(read(a, 'materia', 'Materia') || ''),
          docenteAusente: String(read(a, 'docenteAusente', 'Docente Ausente', 'docente') || '').trim(),
          estadoCobertura: (read(a, 'estadoCobertura', 'Estado Cobertura') || 'Sin cubrir') as CoberturaStatus,
          docenteDecreto: String(read(a, 'docenteDecreto', 'Docente Decreto', 'PF Cubre por Decreto / Observación') || 'Sin cubrir').trim()
        }))
        setAusencias(normAus)
      }

      if (payload.preceptores && Array.isArray(payload.preceptores)) {
        const normPrec = payload.preceptores.map((p: any) => ({
          curso: String(read(p, 'curso', 'Curso') || ''),
          preceptorTitular: String(read(p, 'preceptorTitular', 'Preceptor Titular') || '').trim(),
          preceptorSuplente: String(read(p, 'preceptorSuplente', 'Preceptor Suplente') || '').trim()
        }))
        if (normPrec.some((p: PreceptorAssignment) => p.preceptorTitular || p.preceptorSuplente)) {
          setPreceptores(normPrec)
        }
      }

      if (payload.paroHistorial && Array.isArray(payload.paroHistorial)) {
        const normParo = payload.paroHistorial.map((p: any, index: number) => ({
          id: String(read(p, 'id', 'ID') || `PARO-${index + 1}`),
          fecha: String(read(p, 'fecha', 'Fecha') || new Date().toISOString().split('T')[0]),
          curso: String(read(p, 'curso', 'Curso') || ''),
          horario: String(read(p, 'horario', 'Horario') || ''),
          materia: String(read(p, 'materia', 'Materia') || ''),
          docente: String(read(p, 'docente', 'Docente') || '').trim(),
          estadoParo: (read(p, 'estadoParo', 'Estado Paro') || 'normal') as Attendance,
          observaciones: String(read(p, 'observaciones', 'Observaciones') || ''),
          tipoPersonal: (read(p, 'tipoPersonal', 'Tipo Personal') || 'docente') as 'docente' | 'preceptor',
          grupo: String(read(p, 'grupo', 'Grupo') || '')
        }))
        setParoHistorial(normParo)
      }

      if (payload.historialSuplencias && Array.isArray(payload.historialSuplencias)) {
        const normSup = payload.historialSuplencias.map((s: any, index: number) => ({
          id: String(read(s, 'id', 'ID') || `HS-${index + 1}`),
          fecha: String(read(s, 'fecha', 'Fecha') || ''),
          curso: String(read(s, 'curso', 'Curso') || ''),
          materia: String(read(s, 'materia', 'Materia') || ''),
          titular: String(read(s, 'titular', 'Titular') || ''),
          suplente: String(read(s, 'suplente', 'Suplente') || ''),
          accion: (read(s, 'accion', 'Acción', 'Accion') || 'Alta') as 'Alta' | 'Baja'
        }))
        setHistorialSuplencias(normSup)
      }

      setSyncMessage(`Conectado exitosamente (${Array.isArray(rows) ? rows.length : courses.length} bloques sincronizados)`)
    } catch { 
      setSyncMessage('No se pudo conectar · usando respaldo local') 
    } finally { 
      setSyncing(false) 
    }
  }

  async function save() { 
    if (blocks.length === 0) { alert('Sincronizando datos... Esperá un segundo antes de guardar.'); return }
    setSyncing(true)
    try { 
      await fetch(spreadsheetApi, { 
        method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, 
        body: JSON.stringify({ action: 'replace', rows: blocks, ausencias, preceptores, paroHistorial, historialSuplencias }) 
      }) 
      setSaved(true); setDirty(0); setSyncMessage('Guardado en Google Drive') 
      setTimeout(() => setSaved(false), 2600) 
    } catch { setSyncMessage('No se pudo guardar en Google Drive') } finally { setSyncing(false) } 
  }

  return (
    <>
      <div className="fixed inset-0 w-full h-[100dvh] -z-10 bg-[url('/fondo.jpg')] bg-cover bg-center bg-no-repeat pointer-events-none opacity-90" />
      <div className="min-h-screen text-slate-950 flex flex-col">
        <header className="border-b border-white/20 bg-white/90 backdrop-blur-md shadow-sm print:hidden">
          <div className="mx-auto flex max-w-[1500px] flex-col sm:flex-row items-center justify-between gap-3 px-4 sm:px-5 py-3">
            <div className="flex items-center gap-3 w-full sm:w-auto">
              <img src={logo} alt="Escudo E.E.S.T. N° 6" className="size-12 sm:size-14 object-contain shrink-0" />
              <div>
                <div className="text-[10px] sm:text-[11px] font-extrabold uppercase tracking-[0.22em] text-[#0b5c46]">E.E.S.T. N° 6</div>
                <h1 className="text-base sm:text-lg font-black tracking-tight text-slate-900">Gestión de Horarios, Preceptores y Paros</h1>
                <p className="text-[11px] sm:text-xs text-slate-600">Banfield · Lomas de Zamora</p>
              </div>
            </div>
            <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
              <Button variant="outline" size="sm" onClick={descargarBaseCompletaExcel} className="bg-white/80"><FileSpreadsheet data-icon="inline-start" />Base Completa (Excel)</Button>
              <Button variant="outline" size="sm" onClick={syncSpreadsheet} disabled={syncing} className="bg-white/80"><Download data-icon="inline-start" />{syncing ? 'Conectando…' : 'Sincronizar'}</Button>
            </div>
          </div>
          <nav className="mx-auto flex max-w-[1500px] gap-1 overflow-x-auto px-4 sm:px-5 scrollbar-none">
            <Tab active={view === 'horario'} onClick={() => setView('horario')} icon={<CalendarDays className="size-4 shrink-0" />} label="Horarios" />
            <Tab active={view === 'docentes'} onClick={() => setView('docentes')} icon={<Users className="size-4 shrink-0" />} label="Docentes" />
            <Tab active={view === 'ausencias'} onClick={() => setView('ausencias')} icon={<UserX className="size-4 shrink-0" />} label="Ausencias" />
            <Tab active={view === 'preceptores'} onClick={() => setView('preceptores')} icon={<UserCheck className="size-4 shrink-0" />} label="Preceptores" />
            <Tab active={view === 'paro'} onClick={() => setView('paro')} icon={<ShieldCheck className="size-4 shrink-0" />} label="Registro de Paro" />
            <Tab active={view === 'suplencias'} onClick={() => setView('suplencias')} icon={<ArrowDownUp className="size-4 shrink-0" />} label="Altas/Bajas" />
          </nav>
        </header>
        
        <main className="mx-auto max-w-[1500px] px-4 sm:px-5 py-4 sm:py-6 flex-1 w-full">
          <div className="mb-6 print:hidden">
            <div className="flex flex-col sm:flex-row flex-wrap items-start sm:items-center justify-between gap-3 bg-white/80 backdrop-blur-md p-4 rounded-2xl shadow-sm border border-white/40">
              <div>
                <p className="mb-1 text-xs font-bold uppercase tracking-[0.15em] text-[#0b5c46]">Panel de control</p>
                <h2 className="text-xl sm:text-2xl font-black tracking-tight text-slate-900">
                  {view === 'horario' ? 'Horario semanal y recreos' : view === 'docentes' ? 'Buscar docentes' : view === 'ausencias' ? 'Base de datos de ausencias y cobertura' : view === 'preceptores' ? 'Asignación de preceptores por curso' : view === 'suplencias' ? 'Registro de Altas y Bajas' : 'Historial y control diario de paro'}
                </h2>
                <p className="mt-1 text-xs sm:text-sm text-slate-600">Actualización activa · {syncMessage}</p>
              </div>
              <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                <Button className="bg-[#0b5c46] text-white hover:bg-[#074030] flex-1 sm:flex-none shadow-md" onClick={save} disabled={syncing || saved}>
                  {saved ? <Check data-icon="inline-start" /> : <Cloud data-icon="inline-start" />}
                  {saved ? 'Guardado' : dirty ? `Guardar (${dirty})` : 'Guardar cambios'}
                </Button>
              </div>
            </div>
          </div>

          {view === 'horario' && <ScheduleView course={course} setCourse={setCourse} filter={filter} setFilter={setFilter} selectedDays={selectedDays} setSelectedDays={setSelectedDays} visible={visible} blocks={blocks} setBlocks={setBlocks} courses={courses} preceptores={preceptores} setEditing={setEditing} onAdd={() => setEditing({ id: crypto.randomUUID(), curso: course, turno: 'Teoría', tipo: 'Teoría', dia: 'Lunes', horario: '07:40–08:40', materia: 'Nueva materia', docente: '', situacionRevista: 'Titular' })} onManageRecreos={() => setShowRecreosModal(true)} />}
          {view === 'docentes' && <TeacherView query={query} setQuery={setQuery} teachers={teachers} blocks={blocks} setEditing={setEditing} ausencias={ausencias} />}
          {view === 'ausencias' && <AusenciasView ausencias={ausencias} setAusencias={setAusencias} setDirty={setDirty} courses={courses} blocks={blocks} />}
          {view === 'preceptores' && <PreceptoresView preceptores={preceptores} setPreceptores={setPreceptores} courses={courses} setDirty={setDirty} />}
          {view === 'paro' && <StrikeView dayBlocks={dayBlocks} courses={courses} paroHistorial={paroHistorial} setParoHistorial={setParoHistorial} setDirty={setDirty} preceptores={preceptores} />}
          {view === 'suplencias' && <SuplenciasView historial={historialSuplencias} setHistorial={setHistorialSuplencias} setDirty={setDirty} />}
        </main>
        
        <footer className="mx-auto flex max-w-[1500px] flex-col sm:flex-row items-center justify-between gap-2 border-t border-white/20 px-4 sm:px-5 py-4 text-xs text-slate-600 print:hidden text-center sm:text-left bg-white/80 backdrop-blur-md mt-auto w-full">
          <span>© 2026 E.E.S.T. N° 6 · Sistema de Gestión Escolar</span>
          <span className="flex items-center gap-1.5"><Zap className="size-3 text-[#388e3c]" />Sincronización en vivo con Drive</span>
        </footer>
        {editing && <Editor block={editing} courses={courses} onClose={() => setEditing(null)} onSave={updateBlock} onDelete={() => { 
          if (editing && editing.suplente) {
            setHistorialSuplencias(prev => [{ id: `HS-${Date.now()}-B`, fecha: new Date().toISOString().split('T')[0], curso: editing.curso, materia: editing.materia, titular: editing.docente, suplente: editing.suplente!.trim().toUpperCase(), accion: 'Baja' }, ...prev]);
          }
          setBlocks((bs: Block[]) => bs.filter((b: Block) => b.id !== editing.id)); 
          setEditing(null); 
          setDirty((n) => n + 1) 
        }} />}
        {showRecreosModal && <RecreosModal course={course} blocks={blocks} setBlocks={setBlocks} onClose={() => setShowRecreosModal(false)} setDirty={setDirty} />}
      </div>
    </>
  )
}

function Tab({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactNode; label: string }) { 
  return <button onClick={onClick} className={`flex items-center gap-2 border-b-2 px-3 py-3 text-sm font-bold transition whitespace-nowrap ${active ? 'border-[#0b5c46] text-[#0b5c46]' : 'border-transparent text-slate-500 hover:text-slate-800'}`}>{icon}<span>{label}</span></button> 
}

function ScheduleView({ course, setCourse, filter, setFilter, selectedDays, setSelectedDays, visible, blocks, setBlocks, courses, preceptores, setEditing, onAdd, onManageRecreos }: any) { 
  const tableRef = useRef<HTMLDivElement>(null)
  const [downloadingImg, setDownloadingImg] = useState(false)
  const timePattern = /(\d{1,2}:\d{2})\s*[-– aA]+\s*(\d{1,2}:\d{2})/i;
  
  const courseBlocks = blocks.filter((b: Block) => b.curso === course);
  const boundariesSet = new Set<string>();
  courseBlocks.forEach((b: Block) => {
    const m = b.horario.match(timePattern);
    if (m) { boundariesSet.add(m[1].padStart(5, '0')); boundariesSet.add(m[2].padStart(5, '0')); }
  });
  const boundaries = Array.from(boundariesSet).sort();

  function timeToMins(t: string) {
    const [h, m] = t.split(':').map(Number); return h * 60 + m;
  }

  const rowHeights = boundaries.slice(0, -1).map((start, i) => {
    const end = boundaries[i + 1];
    return `minmax(${Math.max((timeToMins(end) - timeToMins(start)) * 1.1, 40)}px, auto)`; 
  });

  const cellGroups: Record<string, Block[]> = {};
  const untimedBlocks: Block[] = [];
  
  visible.forEach((b: Block) => {
    if (b.tipo === 'Recreo') return;
    const m = b.horario.match(timePattern);
    if (m) {
      const key = `${b.dia}_${m[1].padStart(5, '0')}_${m[2].padStart(5, '0')}`;
      if (!cellGroups[key]) cellGroups[key] = [];
      cellGroups[key].push(b);
    } else untimedBlocks.push(b);
  });

  const recreos = blocks.filter((b: Block) => b.curso === course && b.tipo === 'Recreo');

  function descargarExcelCursoEstetico() {
    const cursoBlocks = blocks.filter((b: Block) => b.curso === course && b.tipo !== 'Recreo')
    const preceptorInfo = preceptores.find((p: PreceptorAssignment) => p.curso === course)
    const wb = XLSX.utils.book_new()
    const exportData = [
      ["E.E.S.T. N° 6 - BANFIELD"],
      [`HORARIO ESCOLAR - CURSO: ${course}`],
      [`Preceptor Titular: ${preceptorInfo?.preceptorTitular || 'No asignado'}  |  Suplente: ${preceptorInfo?.preceptorSuplente || 'Sin asignar'}`],
      [""], ["HORARIO", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]
    ]
    
    boundaries.slice(0, -1).forEach((start, i) => {
      const end = boundaries[i+1];
      const timeStr = `${start}-${end}`;
      const isRecreo = recreos.some((r: Block) => {
        const m = r.horario.match(timePattern);
        return m && m[1].padStart(5, '0') === start && m[2].padStart(5, '0') === end;
      });

      if (isRecreo) {
        exportData.push([timeStr, "RECREO", "RECREO", "RECREO", "RECREO", "RECREO"])
      } else {
        const rowData = [timeStr]
        days.forEach(day => {
          const blocksInCell = cursoBlocks.filter((b: Block) => {
            const m = b.horario.match(timePattern);
            if (!m) return false;
            return b.dia === day && timeToMins(m[1].padStart(5, '0')) <= timeToMins(start) && timeToMins(m[2].padStart(5, '0')) >= timeToMins(end);
          });
          if (blocksInCell.length === 0) rowData.push("")
          else rowData.push(blocksInCell.map((b: Block) => `${b.materia} (${b.docente}${b.suplente ? ' / Sup: ' + b.suplente : ''})${b.grupo ? ' [' + b.grupo + ']' : ''}`).join(" | "))
        })
        exportData.push(rowData)
      }
    })
    
    const ws = XLSX.utils.aoa_to_sheet(exportData)
    ws['!cols'] = [{ wch: 15 }, { wch: 28 }, { wch: 28 }, { wch: 28 }, { wch: 28 }, { wch: 28 }]
    XLSX.utils.book_append_sheet(wb, ws, `Curso_${course.replace('°', '_').replace('°', '')}`)
    XLSX.writeFile(wb, `horario-curso-${course.replace('°', '').replace('°', '')}.xlsx`)
  }

  async function descargarImagenCurso() {
    if (!tableRef.current) return
    setDownloadingImg(true)
    try {
      const dataUrl = await toPng(tableRef.current, { cacheBust: true, quality: 1, pixelRatio: 2, backgroundColor: '#ffffff' })
      const link = document.createElement('a'); link.download = `horario-curso-${course.replace('°', '').replace('°', '')}.png`; link.href = dataUrl; link.click()
    } catch { alert('No se pudo generar la imagen.') } finally { setDownloadingImg(false) }
  }

  const preceptorDelCurso = preceptores.find((p: PreceptorAssignment) => p.curso === course)

  return <>
    <div className="mb-4 flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-3 rounded-2xl border border-[#d7e8c9] bg-white/90 backdrop-blur-sm p-3 shadow-sm print:hidden">
      <label className="flex items-center justify-between sm:justify-start gap-2 text-sm font-bold">Curso 
        <select value={course} onChange={(e) => setCourse(e.target.value)} className="rounded-lg border border-slate-300 bg-white px-3 py-2 font-black outline-none focus:ring-2 focus:ring-[#0b5c46]">
          {courses.map((c: string) => <option key={c}>{c}</option>)}
        </select>
      </label>
      <div className="flex flex-wrap items-center gap-1 rounded-lg bg-slate-100 p-1 justify-center">
        {days.map((day) => <button key={day} onClick={() => setSelectedDays((current: string[]) => current.includes(day) ? current.filter((item) => item !== day) : [...current, day])} className={`rounded-md px-2 py-1.5 text-xs font-bold ${selectedDays.includes(day) ? 'bg-white text-[#0b5c46] shadow-sm' : 'text-slate-400'}`}>{day.slice(0, 3)}</button>)}
      </div>
      <div className="flex rounded-lg bg-slate-100 p-1 justify-center">
        {(['Todos', 'Teoría', 'Taller'] as const).map((x) => <button key={x} onClick={() => setFilter(x)} className={`rounded-md px-3 py-1.5 text-xs font-bold ${filter === x ? 'bg-white text-[#168000] shadow-sm' : 'text-slate-500'}`}>{x === 'Todos' ? 'Todos' : `Solo ${x}`}</button>)}
      </div>
      <div className="flex flex-wrap gap-2 sm:ml-auto w-full sm:w-auto">
        <Button variant="outline" className="flex-1 sm:flex-none border-[#f39c12] text-[#d35400] hover:bg-[#fef9e7] bg-white/80" onClick={onManageRecreos}><Clock className="size-4 mr-1" /> Recreos</Button>
        <Button variant="outline" className="flex-1 sm:flex-none border-blue-600 text-blue-700 hover:bg-blue-50 bg-white/80" onClick={descargarExcelCursoEstetico}><FileSpreadsheet className="size-4 mr-1" /> Excel</Button>
        <Button variant="outline" className="flex-1 sm:flex-none border-purple-600 text-purple-700 hover:bg-purple-50 bg-white/80" onClick={descargarImagenCurso} disabled={downloadingImg}>
          <ImageIcon className="size-4 mr-1" /> {downloadingImg ? 'Generando...' : 'Descargar PNG'}
        </Button>
        <Button className="bg-[#0b5c46] text-white hover:bg-[#074030] flex-1 sm:flex-none shadow-md" onClick={onAdd}><span className="text-lg leading-none">+</span> Agregar</Button>
      </div>
    </div>
    
    <div className="mb-4 flex flex-col sm:flex-row flex-wrap gap-2 text-xs font-semibold justify-between items-start sm:items-center print:hidden bg-white/80 p-2 rounded-xl border border-white/40">
      <div className="flex flex-wrap gap-2">
        <Legend color="bg-[#0052cc]" label="Titular" />
        <Legend color="bg-[#cc0000]" label="Provisional" />
        <Legend color="bg-slate-900" label="Titular interino" />
        <Legend color="bg-[#0b5c46]" label="Suplente" />
      </div>
      <span className="text-slate-600">Doble clic en los cuadrados para editarlos</span>
    </div>

    <div ref={tableRef} className="relative w-full max-w-full bg-white p-4 sm:p-6 rounded-2xl shadow-xl print:shadow-none print:p-0">
      <div className="relative z-10 w-full min-w-[850px]">
        <div className="flex items-center justify-between mb-6 sm:mb-8 px-2">
          <div className="flex items-center gap-4">
            <img src={logo} alt="Escudo E.E.S.T. N° 6" className="w-14 h-14 sm:w-16 sm:h-16 object-contain drop-shadow-sm" />
            <div className="flex flex-col">
              <span className="text-xs sm:text-sm font-extrabold uppercase tracking-[0.25em] text-[#0b5c46]">E.E.S.T. N° 6 · Banfield</span>
              <div className="flex items-baseline gap-2 sm:gap-3">
                <h2 className="text-3xl sm:text-4xl font-black tracking-tighter text-[#0b5c46]">Horario Escolar</h2>
                <span className="text-[#0b5c46] text-lg sm:text-xl font-bold tracking-widest uppercase">CURSO {course}</span>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-4 sm:gap-6">
            {preceptorDelCurso && (
              <div className="flex items-center gap-3">
                <div className="bg-transparent p-1"><UserCheck className="size-6 text-[#0b5c46] opacity-80" /></div>
                <div className="flex flex-col">
                  <span className="text-[10px] font-bold uppercase text-slate-500 tracking-widest leading-none mb-1">Preceptor Titular</span>
                  <span className="text-sm sm:text-base font-black text-slate-800 leading-none uppercase">{preceptorDelCurso.preceptorTitular || 'No asignado'}</span>
                  {preceptorDelCurso.preceptorSuplente && <span className="text-[10px] font-bold text-slate-500 mt-1 uppercase">Sup: {preceptorDelCurso.preceptorSuplente}</span>}
                </div>
              </div>
            )}
            <div className="hidden sm:block border-l border-slate-300 pl-4 py-1">
              <p className="text-[11px] font-semibold italic text-slate-600 leading-tight w-28">Formación técnica,<br />construyendo futuros</p>
            </div>
          </div>
        </div>

        <div className="pb-2">
          {boundaries.length > 0 ? (
            <div style={{ display: 'grid', gridTemplateColumns: `70px repeat(${selectedDays.length}, minmax(130px, 1fr))`, gridTemplateRows: `36px ${rowHeights.join(' ')}` }} className="border border-slate-200 bg-white rounded-xl overflow-hidden relative shadow-sm">
              <div className="border-r border-b border-white/20 bg-[#0b5c46] p-2 flex items-center justify-center font-black text-white text-[10px] uppercase tracking-widest shadow-inner z-30"><Clock className="size-4 mr-1.5 opacity-80" /> HORARIO</div>
              {selectedDays.map((d, i) => <div key={d} style={{ gridColumn: i + 2, gridRow: 1 }} className="border-r border-b border-white/20 bg-[#0b5c46] p-2 flex items-center justify-center font-black text-white text-xs uppercase tracking-[0.2em] shadow-inner z-30">{d}</div>)}
              {selectedDays.map((day, dIdx) => boundaries.slice(0, -1).map((_, rIdx) => <div key={`cell-${dIdx}-${rIdx}`} style={{ gridColumn: dIdx + 2, gridRow: rIdx + 2 }} className="border-r border-b border-slate-100" />))}

              {recreos.map(r => {
                 const m = r.horario.match(timePattern);
                 if (!m || boundaries.indexOf(m[1].padStart(5, '0')) === -1) return null;
                 return (
                   <div key={r.id} style={{ gridColumn: `2 / -1`, gridRow: `${boundaries.indexOf(m[1].padStart(5, '0')) + 2} / ${boundaries.indexOf(m[2].padStart(5, '0')) + 2}` }} onDoubleClick={() => setEditing(r)} className="bg-[#fef9e7] z-0 flex items-center justify-center border-b border-t border-slate-100 cursor-pointer">
                     <div className="flex items-center gap-2 text-[#d35400] opacity-90"><Coffee className="size-4" /><span className="text-[11px] sm:text-xs font-black uppercase tracking-[0.5em]">RECREO</span></div>
                   </div>
                 )
              })}

              {boundaries.slice(0, -1).map((start, rIdx) => {
                const end = boundaries[rIdx + 1];
                const isExplicitRecreo = recreos.some(r => r.horario.match(timePattern)?.[1].padStart(5, '0') === start && r.horario.match(timePattern)?.[2].padStart(5, '0') === end);
                return (
                  <div key={`time-${rIdx}`} style={{ gridColumn: 1, gridRow: rIdx + 2 }} className={`border-r border-b border-slate-200 flex flex-col items-center justify-center p-1 font-mono z-20 ${isExplicitRecreo ? 'bg-[#fef9e7]' : 'bg-[#f8fafc]'}`}>
                    {isExplicitRecreo ? (
                      <><span className="text-[#d35400] text-[11px] sm:text-[13px] font-black leading-none">{start}</span><span className="text-[#e67e22] font-bold text-[8px] sm:text-[9px] leading-none mt-1">{end}</span></>
                    ) : (
                      <><span className="text-slate-800 text-[11px] sm:text-[13px] font-black leading-none">{start}</span><span className="text-slate-400 font-bold text-[8px] sm:text-[9px] mt-1 leading-none">{end}</span></>
                    )}
                  </div>
                )
              })}

              {Object.entries(cellGroups).map(([key, blocksInCell]) => {
                const [dia, start, end] = key.split('_');
                const colIdx = selectedDays.indexOf(dia) + 2;
                if (boundaries.indexOf(start) === -1 || colIdx < 2) return null;
                return (
                  <div key={key} style={{ gridColumn: colIdx, gridRow: `${boundaries.indexOf(start) + 2} / ${boundaries.indexOf(end) + 2}` }} className="p-0.5 sm:p-1 z-10 relative">
                     <div className={`flex gap-1 h-full w-full ${blocksInCell.length > 1 ? 'flex-row' : 'flex-col'}`}>
                        {blocksInCell.map((b: Block) => <div key={b.id} className="flex-1 min-w-0 h-full w-full"><BlockCard block={b} onDoubleClick={() => setEditing(b)} /></div>)}
                     </div>
                  </div>
                )
              })}
            </div>
          ) : (
            <div className="p-10 text-center text-slate-500 font-semibold border-2 border-dashed border-slate-300 rounded-xl bg-slate-50">No hay horarios cargados para este curso.</div>
          )}
        </div>
      </div>
    </div>
  </> 
}

function BlockCard({ block, onDoubleClick }: { block: Block; onDoubleClick: () => void }) { 
  const statusColorClass = block.situacionRevista === 'Titular' ? 'text-[#0052cc]' : block.situacionRevista === 'Provisional' ? 'text-[#cc0000]' : 'text-slate-900'
  return <div onDoubleClick={onDoubleClick} className={`group cursor-pointer rounded-lg border p-1.5 flex flex-col h-full w-full relative shadow-sm hover:shadow-md transition border-l-[4px] sm:border-l-[6px] min-h-max overflow-y-auto overflow-x-hidden scrollbar-none ${block.tipo === 'Teoría' ? 'border-[#8ebf42]/40 border-l-[#0b5c46] bg-[#f4fbe8]' : 'border-slate-300 border-l-slate-400 bg-slate-100'}`}>
    <div className="flex-1 flex flex-col justify-between">
      <div className="flex flex-col gap-0.5 items-start w-full">
        {block.grupo && <span className="rounded bg-[#0b5c46] px-1 py-0.5 text-[7px] sm:text-[8px] font-black text-white shrink-0 uppercase mb-0.5 shadow-sm leading-none">{block.grupo}</span>}
        <p className={`text-[8.5px] sm:text-[11px] font-black uppercase tracking-tight break-words whitespace-normal w-full leading-[1.1] ${block.tipo === 'Teoría' ? 'text-[#0b5c46]' : 'text-[#0b5c46]'}`}>{block.materia}</p>
      </div>
      <div className="mt-1 pt-0.5">
        <div className="flex flex-col">
          {block.docente ? <span className={`${statusColorClass} text-[9px] sm:text-[11px] font-black uppercase leading-[1.1] break-words`}>{block.docente}</span> : <span className="text-[9px] text-slate-400 font-bold uppercase">Sin docente</span>}
          <span className="text-[7.5px] sm:text-[9px] font-mono font-bold text-slate-500 mt-0.5">{block.horario}</span>
        </div>
      </div>
    </div>
    {block.suplente && (
      <div className="mt-0.5 w-fit rounded bg-white px-1 py-0.5 text-[8px] sm:text-[10px] font-black text-[#1b5e20] border border-[#c8e6c9] shrink-0 break-words uppercase whitespace-normal leading-tight">
        SUP: {block.suplente}
      </div>
    )}
  </div> 
}

function Legend({ color, label }: { color: string; label: string }) { 
  return <span className="inline-flex items-center gap-1.5 rounded-full border border-slate-200 bg-white/90 px-2 py-1 shadow-sm"><span className={`size-2 rounded-full ${color}`} />{label}</span> 
}

function TeacherView({ query, setQuery, teachers, blocks, setEditing, ausencias }: any) { 
  const [status, setStatus] = useState('Todos')
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc')
  
  const isExact = teachers.includes(query)
  
  const filtered = isExact 
    ? blocks.filter((b: Block) => {
        const match = b.docente === query || b.suplente === query;
        if (!match) return false;
        if (status === 'Todos') return true;
        if (status === 'Suplentes') return b.suplente === query;
        return b.situacionRevista === status && b.docente === query;
      })
    : []
  
  const selected = isExact ? query : (query ? 'Seleccioná un docente de la lista lateral' : 'Todos los docentes')

  const ausenciasDelDocente = isExact
    ? ausencias
        .filter((a: AusenciaRecord) => a.docenteAusente === query || a.docenteDecreto === query)
        .sort((a: AusenciaRecord, b: AusenciaRecord) => {
          if (sortOrder === 'desc') return (b.fecha || '').localeCompare(a.fecha || '')
          return (a.fecha || '').localeCompare(b.fecha || '')
        })
    : []

  return <div className="grid gap-4 lg:grid-cols-[340px_1fr]">
    <aside className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm p-4 shadow-sm h-fit">
      <p className="mb-2 text-sm font-black text-[#0b5c46]">Buscar docente</p>
      <div className="relative">
        <Search className="absolute left-3 top-2.5 size-4 text-slate-400" />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Nombre o apellido..." className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm outline-none focus:ring-2 focus:ring-[#0b5c46]" />
      </div>
      <div className="mt-3 flex flex-col gap-1 max-h-[300px] sm:max-h-[400px] overflow-y-auto pr-1">
        {teachers.filter((t: string) => t.toLowerCase().includes(query.toLowerCase())).map((t: string) => (
          <button key={t} onClick={() => setQuery(t)} className="flex items-center justify-between rounded-lg px-2 py-2 text-left text-sm font-semibold hover:bg-[#f4faed]">{t}<ChevronDown className="size-3 -rotate-90 text-slate-400" /></button>
        ))}
      </div>
    </aside>
    <section className="flex flex-col gap-4">
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
        <Stat label="Bloques semanales" value={filtered.length} icon={<CalendarDays />} />
        <Stat label="Cursos asignados" value={new Set(filtered.map((b: Block) => b.curso)).size} icon={<Users />} />
        <Stat label="Como titular" value={filtered.filter((b: Block) => b.situacionRevista === 'Titular' && b.docente === query).length} icon={<ShieldCheck />} />
        <Stat label="Como suplente" value={filtered.filter((b: Block) => b.suplente && b.suplente === query).length} icon={<Check />} />
      </div>
      
      <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm shadow-sm overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b p-4">
          <div>
            <h3 className="font-black text-lg text-[#0b5c46]">{selected}</h3>
            <p className="text-xs text-slate-500">
              {isExact ? `${filtered.length} bloques encontrados en el horario semanal` : 'Buscá y hacé clic en un nombre de la lista lateral'}
            </p>
          </div>
          {isExact && (
            <div className="flex flex-wrap gap-1">
              {['Todos', 'Titular', 'Provisional', 'Titular Interino', 'Suplentes'].map((x) => (
                <button key={x} onClick={() => setStatus(x)} className={`rounded-full px-2.5 py-1 text-[11px] font-bold ${status === x ? 'bg-[#0b5c46] text-white shadow-md' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>{x}</button>
              ))}
            </div>
          )}
        </div>
        
        {isExact && (
          <div className="overflow-x-auto">
            <table className="w-full min-w-[720px] text-left text-sm">
              <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
                <tr><th className="p-3">Curso</th><th>Materia</th><th>Día / horario</th><th>Tipo</th><th>Situación</th><th /></tr>
              </thead>
              <tbody>
                {filtered.map((b: Block) => (
                  <tr key={b.id} onDoubleClick={() => setEditing(b)} className="border-t border-slate-100 hover:bg-[#f4faed] cursor-pointer transition">
                    <td className="p-3 font-bold">{b.curso} {b.grupo && <span className="text-xs font-normal text-slate-500">({b.grupo})</span>}</td>
                    <td className="font-bold text-[#1b5014]">{b.materia}<div className="text-xs text-slate-600 font-semibold">{b.docente}</div></td>
                    <td>{b.dia}<div className="font-mono text-xs text-slate-500">{b.horario}</div></td>
                    <td><span className="rounded-full bg-slate-100 px-2 py-1 text-xs font-bold">{b.tipo}</span></td>
                    <td>
                      <span className={`text-xs font-bold ${b.suplente === query ? 'text-green-700' : (statusStyle[b.situacionRevista] || 'text-slate-600')}`}>
                        {b.suplente === query ? 'Suplente' : b.situacionRevista}
                      </span>
                    </td>
                    <td><Pencil className="size-4 text-slate-400" /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {isExact && ausenciasDelDocente.length > 0 && (
        <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm shadow-sm overflow-hidden mt-2">
          <div className="border-b border-white/20 p-4 bg-slate-50/80 flex flex-wrap gap-4 justify-between items-center">
            <div>
              <h3 className="font-black text-base text-[#0b5c46]">Ausencias y Coberturas Registradas</h3>
              <p className="text-xs text-slate-500">Historial de faltas o suplencias por decreto de este docente.</p>
            </div>
            <Button variant="outline" size="sm" onClick={() => setSortOrder(o => o === 'desc' ? 'asc' : 'desc')} className="bg-white">
              <ArrowDownUp className="size-4 mr-2" />
              {sortOrder === 'desc' ? 'Más recientes primero' : 'Más antiguas primero'}
            </Button>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[720px] text-left text-sm">
              <thead className="bg-slate-100/80 text-xs uppercase tracking-wider text-slate-700">
                <tr><th className="p-3">Fecha</th><th>Curso</th><th>Materia</th><th>Rol en el registro</th><th>Estado de cobertura</th></tr>
              </thead>
              <tbody>
                {ausenciasDelDocente.map((a: AusenciaRecord) => {
                  const isAusente = a.docenteAusente === query;
                  return (
                    <tr key={a.id} className={`border-t transition ${isAusente ? 'border-red-100/50 bg-red-50/30 hover:bg-red-50' : 'border-green-100/50 bg-green-50/40 hover:bg-green-100'}`}>
                      <td className="p-3 font-bold">{a.fecha}</td>
                      <td className={`font-bold ${isAusente ? 'text-[#0b5c46]' : 'text-green-800'}`}>{a.curso}</td>
                      <td>{a.materia}</td>
                      <td>
                        {isAusente ? (
                          <span className="text-red-600 font-bold bg-white px-2 py-1 rounded-full text-xs border border-red-200 shadow-sm">Falta</span>
                        ) : (
                          <span className="text-green-700 font-bold bg-white px-2 py-1 rounded-full text-xs border border-green-300 shadow-sm">Cubre por Decreto</span>
                        )}
                      </td>
                      <td>
                        {isAusente ? (
                          <span className={`text-xs font-bold ${a.estadoCobertura === 'Cubierto por Decreto' ? 'text-green-700' : 'text-red-700'}`}>{a.estadoCobertura}</span>
                        ) : (
                          <span className="text-xs font-bold text-slate-500">Reemplaza a {a.docenteAusente}</span>
                        )}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  </div> 
}

function AusenciasView({ ausencias, setAusencias, setDirty, courses, blocks }: any) {
  const [cursoFiltro, setCursoFiltro] = useState('Todos')
  const [docenteFiltro, setDocenteFiltro] = useState('')
  const [modoFecha, setModoFecha] = useState<'todo' | 'rango'>('todo')
  const [fechaDesde, setFechaDesde] = useState('2026-04-01')
  const [fechaHasta, setFechaHasta] = useState('2026-12-31')
  
  const [viewMode, setViewMode] = useState<'lista' | 'calendario'>('lista')
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc')
  const [calendarDate, setCalendarDate] = useState(new Date())

  const [fechaInput, setFechaInput] = useState(new Date().toISOString().split('T')[0])
  const [claseSeleccionada, setClaseSeleccionada] = useState<string>('manual')
  const [docenteInput, setDocenteInput] = useState('')
  const [materiaInput, setMateriaInput] = useState('')
  const [estadoCob, setEstadoCob] = useState<CoberturaStatus>('Sin cubrir')
  const [docenteDecreto, setDocenteDecreto] = useState('')

  const nombreDiaSemana = useMemo(() => {
    const d = new Date(fechaInput + 'T00:00:00')
    const days = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado']
    return days[d.getDay()] || 'Lunes'
  }, [fechaInput])

  const opcionesDelDia = useMemo(() => {
    return blocks.filter((b: Block) => b.curso === (cursoFiltro === 'Todos' ? courses[0] : cursoFiltro) && b.dia === nombreDiaSemana && b.tipo !== 'Recreo')
  }, [blocks, cursoFiltro, nombreDiaSemana, courses])

  useEffect(() => {
    if (claseSeleccionada !== 'manual') {
      const block = blocks.find((b: Block) => b.id === claseSeleccionada)
      if (block) {
        setDocenteInput(block.suplente || block.docente)
        setMateriaInput(block.materia)
      }
    }
  }, [claseSeleccionada, blocks])

  const ausenciasFiltradas = ausencias.filter((a: AusenciaRecord) => {
    const matchCurso = cursoFiltro === 'Todos' || a.curso === cursoFiltro
    const matchDocente = !docenteFiltro || (a.docenteAusente || '').toLowerCase().includes(docenteFiltro.toLowerCase()) || (a.docenteDecreto || '').toLowerCase().includes(docenteFiltro.toLowerCase())
    let matchFecha = true
    if (modoFecha === 'rango' && a.fecha && a.fecha.match(/^\d{4}-\d{2}-\d{2}$/)) {
      matchFecha = a.fecha >= fechaDesde && a.fecha <= fechaHasta
    }
    return matchCurso && matchDocente && matchFecha
  })

  const ausenciasParaTabla = [...ausenciasFiltradas].sort((a: AusenciaRecord, b: AusenciaRecord) => {
    if (sortOrder === 'desc') return (b.fecha || '').localeCompare(a.fecha || '')
    return (a.fecha || '').localeCompare(b.fecha || '')
  })

  const totalAusencias = ausenciasFiltradas.length
  const cubiertas = ausenciasFiltradas.filter((a: AusenciaRecord) => a.estadoCobertura === 'Cubierto por Decreto').length

  function registrarAusencia(e: React.FormEvent) {
    e.preventDefault()
    if (!docenteInput || !cursoFiltro) return
    const nueva: AusenciaRecord = {
      id: `AUS-${Date.now().toString().slice(-4)}`,
      curso: cursoFiltro === 'Todos' ? courses[0] : cursoFiltro,
      mes: 'Octubre', fecha: fechaInput, materia: materiaInput || 'Materia general',
      docenteAusente: docenteInput.toUpperCase(), estadoCobertura: estadoCob,
      docenteDecreto: estadoCob === 'Cubierto por Decreto' ? (docenteDecreto.toUpperCase() || 'CUBIERTO') : 'Sin cubrir'
    }
    setAusencias([nueva, ...ausencias])
    setDirty((n: number) => n + 1)
    setDocenteInput(''); setDocenteDecreto(''); setClaseSeleccionada('manual')
  }

  const calYear = calendarDate.getFullYear()
  const calMonth = calendarDate.getMonth()
  const monthNames = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
  const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate()
  const firstDayRaw = new Date(calYear, calMonth, 1).getDay()
  const firstDay = firstDayRaw === 0 ? 6 : firstDayRaw - 1

  function prevMonth() { setCalendarDate(new Date(calYear, calMonth - 1, 1)) }
  function nextMonth() { setCalendarDate(new Date(calYear, calMonth + 1, 1)) }

  return <div className="flex flex-col gap-6">
    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
      <Stat label="Ausencias en la selección" value={totalAusencias} icon={<UserX />} />
      <Stat label="Cubiertas por decreto" value={cubiertas} icon={<ShieldCheck />} />
      <Stat label="Sin cubrir" value={totalAusencias - cubiertas} icon={<ShieldAlert />} />
    </div>

    <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm p-4 sm:p-5 shadow-sm">
      <h3 className="mb-3 font-black text-base flex items-center gap-2 text-[#0b5c46]"><Search className="size-4" /> Filtrar y buscar ausencias</h3>
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Filtrar por Curso
          <select value={cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold">
            <option value="Todos">Todos los cursos</option>
            {courses.map((c: string) => <option key={c}>{c}</option>)}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Buscar por Docente
          <input value={docenteFiltro} onChange={(e) => setDocenteFiltro(e.target.value)} placeholder="Ej. BOGARIN, ASPREA..." className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Periodo de fechas
          <select value={modoFecha} onChange={(e) => setModoFecha(e.target.value as any)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold">
            <option value="todo">Todo el historial</option><option value="rango">Desde / Hasta específico</option>
          </select>
        </label>
        {modoFecha === 'rango' && (
          <div className="grid grid-cols-2 gap-2">
            <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-600">Desde <input type="date" value={fechaDesde} onChange={(e) => setFechaDesde(e.target.value)} className="rounded-lg border border-slate-200 px-2 py-2 text-xs" /></label>
            <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-600">Hasta <input type="date" value={fechaHasta} onChange={(e) => setFechaHasta(e.target.value)} className="rounded-lg border border-slate-200 px-2 py-2 text-xs" /></label>
          </div>
        )}
      </div>
    </div>

    <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm p-4 sm:p-5 shadow-sm">
      <h3 className="mb-4 font-black text-lg text-[#0b5c46]">Registrar nueva ausencia docente</h3>
      <form onSubmit={registrarAusencia} className="grid gap-4 sm:grid-cols-2 lg:grid-cols-6">
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Curso
          <select value={cursoFiltro === 'Todos' ? (courses[0] || '1°1°') : cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-[#0b5c46]">
            {courses.map((c: string) => <option key={c}>{c}</option>)}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Fecha
          <input type="date" value={fechaInput} onChange={(e) => setFechaInput(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-[#0b5c46]" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600 sm:col-span-2 lg:col-span-4">Clase programada ({nombreDiaSemana})
          <select value={claseSeleccionada} onChange={(e) => setClaseSeleccionada(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-bold text-[#0b5c46] outline-none focus:ring-2 focus:ring-[#0b5c46] bg-slate-50">
            <option value="manual">-- Ingreso manual libre --</option>
            {opcionesDelDia.map((b: Block) => <option key={b.id} value={b.id}>{b.horario} hs | {b.materia} ({b.suplente || b.docente})</option>)}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600 sm:col-span-1 lg:col-span-2">Docente ausente
          <input value={docenteInput} onChange={(e) => setDocenteInput(e.target.value)} placeholder="Ej. BOGARIN" className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase" required />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600 sm:col-span-1 lg:col-span-2">Materia
          <input value={materiaInput} onChange={(e) => setMateriaInput(e.target.value)} placeholder="Ej. Matemática" className="rounded-lg border border-slate-200 px-3 py-2 text-sm" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600 sm:col-span-1 lg:col-span-2">Estado de Cobertura
          <select value={estadoCob} onChange={(e) => setEstadoCob(e.target.value as CoberturaStatus)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold outline-none focus:ring-2 focus:ring-[#0b5c46]">
            <option value="Sin cubrir">Sin cubrir</option><option value="Cubierto por Decreto">Cubierto por Decreto</option>
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600 sm:col-span-1 lg:col-span-4">Profesor que cubre / Obs.
          <input value={docenteDecreto} onChange={(e) => setDocenteDecreto(e.target.value)} placeholder="Ej. ASPREA" disabled={estadoCob === 'Sin cubrir'} className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase disabled:bg-slate-100" />
        </label>
        <div className="sm:col-span-2 lg:col-span-2 flex items-end justify-end">
          <Button type="submit" className="bg-[#0b5c46] text-white hover:bg-[#074030] w-full shadow-md h-[38px]">Registrar ausencia</Button>
        </div>
      </form>
    </div>

    <div className="overflow-hidden rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm shadow-sm flex flex-col">
      <div className="border-b border-white/20 p-4 flex flex-col sm:flex-row gap-4 items-start sm:items-center justify-between">
        <div>
          <h3 className="font-black text-lg text-[#0b5c46]">Historial de ausencias y cobertura</h3>
          <p className="text-xs text-slate-500">Mostrando {ausenciasFiltradas.length} registros según los filtros seleccionados</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {viewMode === 'lista' && (
            <Button variant="outline" size="sm" onClick={() => setSortOrder(o => o === 'desc' ? 'asc' : 'desc')} className="bg-white">
              <ArrowDownUp className="size-4 mr-1.5" />
              {sortOrder === 'desc' ? 'Más recientes' : 'Más antiguas'}
            </Button>
          )}
          <div className="bg-slate-100 p-1 rounded-lg flex border border-slate-200">
            <button onClick={() => setViewMode('lista')} className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 ${viewMode === 'lista' ? 'bg-white shadow-sm text-[#0b5c46]' : 'text-slate-500 hover:text-slate-700'}`}>
              <ListFilter className="size-3.5" /> Lista
            </button>
            <button onClick={() => setViewMode('calendario')} className={`px-3 py-1.5 text-xs font-bold rounded-md flex items-center gap-1.5 ${viewMode === 'calendario' ? 'bg-white shadow-sm text-[#0b5c46]' : 'text-slate-500 hover:text-slate-700'}`}>
              <CalendarDays className="size-3.5" /> Calendario
            </button>
          </div>
        </div>
      </div>

      {viewMode === 'lista' ? (
        <div className="overflow-x-auto">
          <table className="w-full min-w-[850px] text-left text-sm">
            <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
              <tr><th className="p-3">Curso</th><th>Fecha</th><th>Docente ausente</th><th>Materia</th><th>Estado</th><th>Cubre por decreto / Observación</th></tr>
            </thead>
            <tbody>
              {ausenciasParaTabla.length === 0 ? (
                <tr><td colSpan={6} className="p-6 text-center text-slate-400 font-semibold">No se encontraron ausencias con los filtros seleccionados.</td></tr>
              ) : (
                ausenciasParaTabla.map((a: AusenciaRecord) => (
                  <tr key={a.id} className="border-t border-slate-100 hover:bg-[#f4faed]">
                    <td className="p-3 font-bold">{a.curso}</td><td>{a.fecha}</td><td className="font-black text-slate-800">{a.docenteAusente}</td><td className="font-bold text-[#1b5014]">{a.materia}</td>
                    <td><span className={`rounded-full px-2.5 py-1 text-xs font-bold shadow-sm border ${a.estadoCobertura === 'Cubierto por Decreto' ? 'bg-green-100 text-green-800 border-green-200' : 'bg-red-50 text-red-700 border-red-100'}`}>{a.estadoCobertura}</span></td>
                    <td className="font-bold text-slate-700">{a.docenteDecreto}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="p-4 sm:p-6 bg-slate-50/50">
          <div className="flex items-center justify-between mb-6">
            <Button variant="outline" onClick={prevMonth} className="bg-white px-3 font-black text-lg">{'<'}</Button>
            <div className="text-center">
              <h2 className="text-xl sm:text-2xl font-black uppercase text-slate-900 tracking-tight">{monthNames[calMonth]}</h2>
              <p className="text-lg font-bold text-red-600 leading-none">{calYear}</p>
            </div>
            <Button variant="outline" onClick={nextMonth} className="bg-white px-3 font-black text-lg">{'>'}</Button>
          </div>
          
          <div className="overflow-x-auto">
            <div className="min-w-[700px] grid grid-cols-7 border border-[#2c3e50] rounded-xl overflow-hidden bg-white shadow-sm">
              {['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'].map(d => (
                <div key={d} className="bg-[#2c3e50] text-white text-center py-2.5 text-xs font-bold tracking-wider">{d}</div>
              ))}
              
              {Array.from({ length: firstDay }).map((_, i) => <div key={`blank-${i}`} className="border-b border-r border-slate-200 bg-slate-50 min-h-[120px]" />)}
              
              {Array.from({ length: daysInMonth }).map((_, i) => {
                const day = i + 1;
                const dateStr = `${calYear}-${String(calMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                const dayAbsences = ausenciasFiltradas.filter(a => a.fecha === dateStr);
                const isToday = dateStr === new Date().toISOString().split('T')[0];

                return (
                  <div key={day} className="border-b border-r border-slate-200 p-2 min-h-[120px] flex flex-col hover:bg-slate-50 transition">
                    <span className={`text-sm font-black mb-1 ${isToday ? 'text-blue-600' : 'text-slate-800'}`}>{day}</span>
                    <div className="flex flex-col gap-1.5 overflow-y-auto flex-1 h-0 scrollbar-none">
                      {dayAbsences.map(a => (
                        <div key={a.id} className="text-[9px] leading-tight bg-red-50 text-red-700 p-1.5 rounded-md border border-red-100 font-bold shadow-sm uppercase flex flex-col gap-0.5">
                          <span className="text-[8px] bg-red-200/50 w-fit px-1 rounded-sm text-red-800">{a.curso}</span>
                          <span className="line-clamp-2">{a.docenteAusente}</span>
                          <span className="text-[7.5px] text-red-600/90 font-semibold line-clamp-1 leading-none">{a.materia}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  </div>
}

function PreceptoresView({ preceptores, setPreceptores, courses, setDirty }: any) {
  const [filtroCurso, setFiltroCurso] = useState('')

  function actualizarPreceptor(curso: string, campo: 'preceptorTitular' | 'preceptorSuplente', valor: string) {
    setPreceptores(preceptores.map((p: PreceptorAssignment) => p.curso === curso ? { ...p, [campo]: valor.toUpperCase() } : p))
    setDirty((n: number) => n + 1)
  }

  const filtrados = preceptores.filter((p: PreceptorAssignment) => !filtroCurso || p.curso.toLowerCase().includes(filtroCurso.toLowerCase()) || p.preceptorTitular.toLowerCase().includes(filtroCurso.toLowerCase()) || p.preceptorSuplente.toLowerCase().includes(filtroCurso.toLowerCase()))

  return <div className="flex flex-col gap-6">
    <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm p-4 sm:p-5 shadow-sm">
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4">
        <div>
          <h3 className="font-black text-lg text-[#0b5c46]">Asignación de Preceptores por Curso</h3>
          <p className="text-xs text-slate-500">Configurá el preceptor titular y suplente asignado a cada división de la escuela</p>
        </div>
        <div className="w-full sm:w-auto">
          <input value={filtroCurso} onChange={(e) => setFiltroCurso(e.target.value)} placeholder="Filtrar curso o preceptor..." className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm w-full sm:w-64 focus:ring-2 focus:ring-[#0b5c46] outline-none" />
        </div>
      </div>
      <div className="overflow-x-auto rounded-xl border border-slate-200">
        <table className="w-full min-w-[600px] text-left text-sm">
          <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
            <tr><th className="p-3">Curso / División</th><th>Preceptor Titular</th><th>Preceptor Suplente</th></tr>
          </thead>
          <tbody>
            {filtrados.map((p: PreceptorAssignment) => (
              <tr key={p.curso} className="border-t border-slate-100 hover:bg-[#f4faed] transition">
                <td className="p-3 font-black text-lg text-[#0b5c46]">{p.curso}</td>
                <td className="p-2">
                  <input value={p.preceptorTitular} onChange={(e) => actualizarPreceptor(p.curso, 'preceptorTitular', e.target.value)} placeholder="Nombre del preceptor titular" className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm uppercase font-bold focus:ring-2 focus:ring-[#0b5c46] outline-none" />
                </td>
                <td className="p-2">
                  <input value={p.preceptorSuplente} onChange={(e) => actualizarPreceptor(p.curso, 'preceptorSuplente', e.target.value)} placeholder="Opcional (Suplente)" className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm uppercase font-bold focus:ring-2 focus:ring-[#0b5c46] outline-none" />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  </div>
}

function StrikeView({ dayBlocks, courses, paroHistorial, setParoHistorial, setDirty, preceptores }: any) { 
  const paroRef = useRef<HTMLDivElement>(null)
  const [downloadingImg, setDownloadingImg] = useState(false)

  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [cursoFiltro, setCursoFiltro] = useState('Todos')
  
  const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado']
  const dateObj = new Date(date + 'T00:00:00')
  const nombreDiaSemana = dayNames[dateObj.getDay()] || 'Lunes'
  const esFinDeSemana = nombreDiaSemana === 'Domingo' || nombreDiaSemana === 'Sábado'

  const bloquesDelDia = useMemo(() => {
    if (esFinDeSemana) return []
    let list = (dayBlocks || []).filter((b: Block) => b && b.dia === nombreDiaSemana && b.tipo !== 'Recreo')
    if (cursoFiltro !== 'Todos') list = list.filter((b: Block) => b.curso === cursoFiltro)
    return list
  }, [dayBlocks, date, cursoFiltro, nombreDiaSemana, esFinDeSemana])

  const preceptoresDelDia = useMemo(() => {
    if (esFinDeSemana) return []
    if (cursoFiltro !== 'Todos') return (preceptores || []).filter((p: PreceptorAssignment) => p.curso === cursoFiltro)
    return preceptores || []
  }, [preceptores, cursoFiltro, esFinDeSemana])

  function cambiarEstadoDocente(b: any, estado: Attendance, obs?: string) {
    const grupoKey = b.grupo || ''
    const existingIndex = paroHistorial.findIndex((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia && (p.grupo || '') === grupoKey && p.tipoPersonal === 'docente')
    const nuevoRegistro: ParoRecord = {
      id: existingIndex >= 0 ? paroHistorial[existingIndex].id : `PARO-D-${Date.now().toString().slice(-4)}`,
      fecha: date, curso: b.curso, horario: b.horario, materia: b.materia,
      docente: b.suplente || b.docente, estadoParo: estado,
      observaciones: obs !== undefined ? obs : (existingIndex >= 0 ? paroHistorial[existingIndex].observaciones : ''),
      tipoPersonal: 'docente', grupo: grupoKey
    }
    if (existingIndex >= 0) {
      const copy = [...paroHistorial]; copy[existingIndex] = nuevoRegistro; setParoHistorial(copy)
    } else setParoHistorial([nuevoRegistro, ...paroHistorial])
    setDirty((n: number) => n + 1)
  }

  function cambiarEstadoPreceptor(p: PreceptorAssignment, estado: Attendance, obs?: string) {
    const existingIndex = paroHistorial.findIndex((item: ParoRecord) => item.fecha === date && item.curso === p.curso && item.tipoPersonal === 'preceptor')
    const nuevoRegistro: ParoRecord = {
      id: existingIndex >= 0 ? paroHistorial[existingIndex].id : `PARO-P-${Date.now().toString().slice(-4)}`,
      fecha: date, curso: p.curso, horario: 'Jornada Completa', materia: 'Preceptoría',
      docente: p.preceptorTitular || 'Preceptor sin nombre', estadoParo: estado,
      observaciones: obs !== undefined ? obs : (existingIndex >= 0 ? paroHistorial[existingIndex].observaciones : ''),
      tipoPersonal: 'preceptor'
    }
    if (existingIndex >= 0) {
      const copy = [...paroHistorial]; copy[existingIndex] = nuevoRegistro; setParoHistorial(copy)
    } else setParoHistorial([nuevoRegistro, ...paroHistorial])
    setDirty((n: number) => n + 1)
  }

  function descargarExcelParo() {
    const wb = XLSX.utils.book_new()
    const exportData = [
      ["E.E.S.T. N° 6 - BANFIELD"], [`REGISTRO DE PARO / ASISTENCIA - Fecha: ${date} (${nombreDiaSemana})`], [`Curso filtrado: ${cursoFiltro}`], [""]
    ]

    if (!esFinDeSemana && preceptoresDelDia.length > 0) {
      exportData.push(["PRECEPTORES"]); exportData.push(["Curso / División", "Preceptor Titular", "Estado", "Observaciones"])
      preceptoresDelDia.forEach((p: PreceptorAssignment) => {
        const reg = paroHistorial.find((item: ParoRecord) => item.fecha === date && item.curso === p.curso && item.tipoPersonal === 'preceptor')
        exportData.push([p.curso, p.preceptorTitular || "Sin asignar", reg?.estadoParo || 'normal', reg?.observaciones || ''])
      }); exportData.push([""])
    }

    exportData.push(["DOCENTES"]); exportData.push(["Curso / Grupo", "Horario", "Materia", "Docente de turno", "Estado", "Observaciones"])
    bloquesDelDia.forEach((b: Block) => {
      const grupoKey = b.grupo || ''
      const reg = paroHistorial.find((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia && (p.grupo || '') === grupoKey && p.tipoPersonal === 'docente')
      exportData.push([`${b.curso}${b.grupo ? ' [' + b.grupo + ']' : ''}`, b.horario, b.materia, b.suplente || b.docente, reg?.estadoParo || 'normal', reg?.observaciones || ''])
    })

    const ws = XLSX.utils.aoa_to_sheet(exportData)
    ws['!cols'] = [{ wch: 15 }, { wch: 15 }, { wch: 30 }, { wch: 30 }, { wch: 15 }, { wch: 40 }]
    XLSX.utils.book_append_sheet(wb, ws, `Paro_${date}`); XLSX.writeFile(wb, `registro-paro-${date}${cursoFiltro !== 'Todos' ? '-' + cursoFiltro : ''}.xlsx`)
  }

  async function descargarImagenParo() {
    if (!paroRef.current) return
    setDownloadingImg(true)
    try {
      const dataUrl = await toPng(paroRef.current, { cacheBust: true, quality: 1, pixelRatio: 2 })
      const link = document.createElement('a'); link.download = `registro-paro-${date}${cursoFiltro !== 'Todos' ? '-' + cursoFiltro : ''}.png`; link.href = dataUrl; link.click()
    } catch { alert('No se pudo generar la imagen.') } finally { setDownloadingImg(false) }
  }

  const totalDocentesParo = bloquesDelDia.filter((b: Block) => {
    const reg = paroHistorial.find((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia && (p.grupo || '') === (b.grupo || '') && p.tipoPersonal === 'docente')
    return reg?.estadoParo === 'paro'
  }).length

  const totalPreceptoresParo = preceptoresDelDia.filter((p: PreceptorAssignment) => {
    const reg = paroHistorial.find((item: ParoRecord) => item.fecha === date && item.curso === p.curso && item.tipoPersonal === 'preceptor')
    return reg?.estadoParo === 'paro'
  }).length

  return <div className="flex flex-col gap-6">
    <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-3 rounded-2xl border border-[#d7e8c9] bg-white/90 backdrop-blur-sm p-3 shadow-sm print:hidden">
      <label className="flex items-center gap-2 text-sm font-bold">Fecha <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold outline-none focus:ring-2 focus:ring-[#0b5c46]" /></label>
      <span className="rounded-full bg-[#0b5c46] text-white px-3 py-2 text-sm font-bold shadow-md">{nombreDiaSemana}</span>
      <label className="flex items-center gap-2 text-sm font-bold ml-0 sm:ml-4">Curso 
        <select value={cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-bold outline-none focus:ring-2 focus:ring-[#0b5c46]">
          <option value="Todos">Todos los cursos</option>
          {courses.map((c: string) => <option key={c}>{c}</option>)}
        </select>
      </label>
      <div className="flex flex-wrap gap-2 sm:ml-auto w-full sm:w-auto">
        <Button variant="outline" className="flex-1 sm:flex-none border-blue-600 text-blue-700 hover:bg-blue-50 bg-white/80" onClick={descargarExcelParo}><FileSpreadsheet className="size-4 mr-1" /> Excel Paro</Button>
        <Button variant="outline" className="flex-1 sm:flex-none border-purple-600 text-purple-700 hover:bg-purple-50 bg-white/80" onClick={descargarImagenParo} disabled={downloadingImg}><ImageIcon className="size-4 mr-1" /> {downloadingImg ? 'Generando...' : 'Descargar Imagen (PNG)'}</Button>
      </div>
    </div>
    
    <div ref={paroRef} className="flex flex-col gap-4 p-4 sm:p-6 rounded-2xl border border-slate-200 bg-white shadow-xl print:p-0 print:border-none print:shadow-none relative overflow-hidden">
      <div className="relative z-10 flex items-center justify-between border-b border-[#0b5c46]/20 pb-4">
        <div className="flex items-center gap-4">
          <img src={logo} alt="Escudo E.E.S.T. N° 6" className="size-14 sm:size-16 object-contain drop-shadow-md" />
          <div>
            <div className="text-[10px] sm:text-xs font-extrabold uppercase tracking-[0.2em] text-[#0b5c46]">E.E.S.T. N° 6 · Banfield</div>
            <h3 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Registro de Paro y Asistencia</h3>
            <p className="text-sm font-bold text-[#0b5c46] mt-0.5">{date} ({nombreDiaSemana})</p>
          </div>
        </div>
        <div className="text-right text-xs font-semibold text-slate-700 bg-slate-50 p-2 rounded-xl border border-slate-100">
          <div>Curso filtrado: <strong className="text-slate-900 text-sm block">{cursoFiltro}</strong></div>
        </div>
      </div>

      <div className="relative z-10 grid grid-cols-1 sm:grid-cols-3 gap-3 print:hidden">
        <Stat label={`Docentes con clase este ${nombreDiaSemana}`} value={bloquesDelDia.length} icon={<ListFilter />} />
        <Stat label="Paros de Docentes" value={totalDocentesParo} icon={<ShieldCheck />} />
        <Stat label="Paros de Preceptores" value={totalPreceptoresParo} icon={<UserCheck />} />
      </div>

      {!esFinDeSemana && preceptoresDelDia.length > 0 && (
        <div className="relative z-10 overflow-hidden rounded-xl border border-[#0b5c46]/30 shadow-sm bg-white">
          <div className="border-b bg-[#0b5c46]/10 p-3"><h3 className="font-black text-sm text-[#0b5c46] flex items-center gap-2"><UserCheck className="size-4"/> Control de Preceptores</h3></div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[750px] text-left text-sm">
              <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
                <tr><th className="p-3">Curso / División</th><th>Preceptor Titular</th><th>Estado (¿Hizo Paro?)</th><th>Observaciones</th></tr>
              </thead>
              <tbody>
                {preceptoresDelDia.map((p: PreceptorAssignment) => {
                  const reg = paroHistorial.find((item: ParoRecord) => item.fecha === date && item.curso === p.curso && item.tipoPersonal === 'preceptor')
                  const estadoActual = reg?.estadoParo || 'normal'
                  return <tr key={p.curso} className="border-t border-slate-100 hover:bg-[#f4faed] transition">
                    <td className="p-3 font-black text-lg text-[#0b5c46]">{p.curso}</td>
                    <td className="font-bold text-slate-800 uppercase">{p.preceptorTitular || <span className="text-slate-400 font-normal">Sin asignar</span>} {p.preceptorSuplente && <span className="block text-xs font-normal text-slate-500">Suplente: {p.preceptorSuplente}</span>}</td>
                    <td>
                      <div className="flex gap-1 print:hidden">
                        <button type="button" onClick={() => cambiarEstadoPreceptor(p, 'normal')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'normal' ? 'bg-green-100 text-green-700 ring-2 ring-green-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Normal</button>
                        <button type="button" onClick={() => cambiarEstadoPreceptor(p, 'paro')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'paro' ? 'bg-red-100 text-red-700 ring-2 ring-red-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Hizo Paro</button>
                        <button type="button" onClick={() => cambiarEstadoPreceptor(p, 'licencia')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'licencia' ? 'bg-amber-100 text-amber-700 ring-2 ring-amber-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Licencia</button>
                        <button type="button" onClick={() => cambiarEstadoPreceptor(p, 'no_contesta')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'no_contesta' ? 'bg-slate-700 text-white ring-2 ring-slate-900 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>No Contesta</button>
                      </div>
                      <div className="hidden print:block font-bold">{estadoActual === 'normal' ? '✅ Normal' : estadoActual === 'paro' ? '🛑 Hizo Paro' : estadoActual === 'licencia' ? '⚠️ Licencia' : '❓ No Contesta'}</div>
                    </td>
                    <td className="p-2"><input value={reg?.observaciones || ''} onChange={(e) => cambiarEstadoPreceptor(p, estadoActual, e.target.value)} placeholder="Observación preceptor..." className="w-full min-w-40 rounded-lg border border-slate-200 px-3 py-2 text-xs outline-none focus:ring-2 focus:ring-[#0b5c46]" /></td>
                  </tr>
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="relative z-10 overflow-hidden rounded-xl border border-[#0b5c46]/30 shadow-sm bg-white">
        <div className="border-b bg-[#0b5c46]/10 p-3"><h3 className="font-black text-sm text-[#0b5c46] flex items-center gap-2"><Users className="size-4"/> Registro de Paro Docente</h3></div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[850px] text-left text-sm">
            <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
              <tr><th className="p-3">Curso / horario</th><th>Materia</th><th>Docente de turno</th><th>Estado (¿Hizo Paro?)</th><th>Observaciones</th></tr>
            </thead>
            <tbody>
              {esFinDeSemana ? (
                <tr><td colSpan={5} className="p-6 text-center text-slate-400 font-semibold">Los fines de semana no hay actividad escolar programada.</td></tr>
              ) : bloquesDelDia.length === 0 ? (
                <tr><td colSpan={5} className="p-6 text-center text-slate-400 font-semibold">No hay docentes programados para este día con los filtros seleccionados.</td></tr>
              ) : (
                bloquesDelDia.map((b: Block) => {
                  const reg = paroHistorial.find((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia && (p.grupo || '') === (b.grupo || '') && p.tipoPersonal === 'docente')
                  const estadoActual = reg?.estadoParo || 'normal'
                  return <tr key={b.id + b.curso + (b.grupo || '')} className="border-t border-slate-100 hover:bg-[#f4faed] transition">
                    <td className="p-3 font-bold">{b.curso} {b.grupo && <span className="rounded bg-[#0b5c46] px-1.5 py-0.5 text-[9px] font-black text-white ml-1 shadow-sm uppercase">{b.grupo}</span>}<div className="font-mono text-xs font-bold text-slate-500 mt-0.5">{b.horario}</div></td>
                    <td className="font-bold text-[#1b5014]">{b.materia}</td>
                    <td>
                      <span className={`uppercase font-black ${b.suplente ? 'text-green-700' : 'text-slate-800'}`}>{b.suplente || b.docente}</span>
                      {b.suplente && <div className="text-[10px] font-bold text-slate-400 uppercase mt-0.5">Titular: {b.docente}</div>}
                    </td>
                    <td>
                      <div className="flex gap-1 print:hidden">
                        <button type="button" onClick={() => cambiarEstadoDocente(b, 'normal')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'normal' ? 'bg-green-100 text-green-700 ring-2 ring-green-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Normal</button>
                        <button type="button" onClick={() => cambiarEstadoDocente(b, 'paro')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'paro' ? 'bg-red-100 text-red-700 ring-2 ring-red-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Hizo Paro</button>
                        <button type="button" onClick={() => cambiarEstadoDocente(b, 'licencia')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'licencia' ? 'bg-amber-100 text-amber-700 ring-2 ring-amber-500 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Licencia</button>
                        <button type="button" onClick={() => cambiarEstadoDocente(b, 'no_contesta')} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${estadoActual === 'no_contesta' ? 'bg-slate-700 text-white ring-2 ring-slate-900 shadow-sm' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>No Contesta</button>
                      </div>
                      <div className="hidden print:block font-bold">{estadoActual === 'normal' ? '✅ Normal' : estadoActual === 'paro' ? '🛑 Hizo Paro' : estadoActual === 'licencia' ? '⚠️ Licencia' : '❓ No Contesta'}</div>
                    </td>
                    <td className="p-2"><input value={reg?.observaciones || ''} onChange={(e) => cambiarEstadoDocente(b, estadoActual, e.target.value)} placeholder="Observación..." className="w-full min-w-40 rounded-lg border border-slate-200 px-3 py-2 text-xs outline-none focus:ring-2 focus:ring-[#0b5c46]" /></td>
                  </tr>
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div> 
}

function SuplenciasView({ historial, setHistorial, setDirty }: any) {
  const [query, setQuery] = useState('')

  const filtrados = (historial || []).filter((h: any) => 
    !query || 
    (h.suplente || '').toLowerCase().includes(query.toLowerCase()) ||
    (h.titular || '').toLowerCase().includes(query.toLowerCase()) ||
    (h.curso || '').toLowerCase().includes(query.toLowerCase()) ||
    (h.materia || '').toLowerCase().includes(query.toLowerCase())
  );

  return (
    <div className="flex flex-col gap-6">
      <div className="rounded-2xl border border-white/40 bg-white/90 backdrop-blur-sm p-4 sm:p-6 shadow-xl">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div>
            <h3 className="font-black text-lg text-[#0b5c46]">Registro de Altas y Bajas de Suplentes</h3>
            <p className="text-sm text-slate-500 mt-1">Este registro automático guarda un historial de cuándo se agregan o quitan suplentes del horario semanal.</p>
          </div>
          <div className="relative w-full sm:w-72 shrink-0">
            <Search className="absolute left-3 top-2.5 size-4 text-slate-400" />
            <input 
              value={query} 
              onChange={(e) => setQuery(e.target.value)} 
              placeholder="Buscar docente o curso..." 
              className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm outline-none focus:ring-2 focus:ring-[#0b5c46]" 
            />
          </div>
        </div>
        
        <div className="overflow-x-auto rounded-xl border border-slate-200">
          <table className="w-full min-w-[700px] text-left text-sm">
            <thead className="bg-[#0b5c46] text-xs uppercase tracking-wider text-white">
              <tr><th className="p-3">Fecha</th><th>Acción</th><th>Suplente</th><th>Titular reemplazado</th><th>Curso</th><th>Materia</th><th/></tr>
            </thead>
            <tbody>
              {filtrados.length === 0 ? (
                <tr><td colSpan={7} className="p-6 text-center text-slate-400 font-semibold">No se encontraron registros con esa búsqueda.</td></tr>
              ) : (
                filtrados.map((h: any) => (
                  <tr key={h.id} className="border-t border-slate-100 hover:bg-[#f4faed] transition">
                    <td className="p-3 font-bold">{h.fecha}</td>
                    <td>
                      <span className={`px-2 py-1 rounded-full text-[10px] font-black uppercase shadow-sm border ${h.accion === 'Alta' ? 'bg-green-100 text-green-800 border-green-200' : 'bg-red-50 text-red-700 border-red-100'}`}>
                        {h.accion}
                      </span>
                    </td>
                    <td className="font-black text-slate-800 uppercase">{h.suplente}</td>
                    <td className="font-semibold text-slate-600 uppercase">{h.titular}</td>
                    <td className="font-bold text-[#0b5c46]">{h.curso}</td>
                    <td className="uppercase text-xs font-bold text-slate-700">{h.materia}</td>
                    <td>
                       <Button variant="ghost" size="sm" onClick={() => { setHistorial(historial.filter((x: any) => x.id !== h.id)); setDirty((n: number) => n + 1) }} className="text-red-600 hover:text-red-700 h-8 px-2"><Trash2 className="size-4" /></Button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function Stat({ label, value, icon }: { label: string; value: number; icon: React.ReactNode }) { 
  return <div className="rounded-2xl border border-[#d7e8c9] bg-white/90 backdrop-blur-sm p-4 shadow-sm"><div className="mb-3 flex items-center justify-between text-[#0b5c46]">{icon}<span className="text-2xl font-black text-slate-900">{value}</span></div><p className="text-xs font-bold text-slate-500 uppercase">{label}</p></div> 
}

function Editor({ block, courses, onClose, onSave, onDelete }: { block: Block; courses: string[]; onClose: () => void; onSave: (b: Block) => void; onDelete: () => void }) { 
  const [draft, setDraft] = useState(block) 
  const set = (key: keyof Block, value: string) => setDraft((d: Block) => ({ ...d, [key]: value })) 
  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/40 p-4 backdrop-blur-sm" role="dialog" aria-modal="true">
    <div className="w-full max-w-lg rounded-2xl bg-white shadow-2xl">
      <div className="flex items-center justify-between border-b p-5">
        <div><p className="text-xs font-bold uppercase tracking-widest text-[#0b5c46]">Edición rápida</p><h2 className="text-xl font-black">{draft.materia}</h2></div>
        <button onClick={onClose} aria-label="Cerrar"><X /></button>
      </div>
      <div className="grid gap-4 p-5 sm:grid-cols-2">
        <Field label="Curso"><select value={draft.curso} onChange={(e) => set('curso', e.target.value)}>{courses.map((c: string) => <option key={c}>{c}</option>)}</select></Field>
        <Field label="Materia"><input value={draft.materia} onChange={(e) => set('materia', e.target.value)} /></Field>
        <Field label="Docente"><input value={draft.docente} onChange={(e) => set('docente', e.target.value)} /></Field>
        <Field label="Día"><select value={draft.dia} onChange={(e) => set('dia', e.target.value)}>{days.map((d) => <option key={d}>{d}</option>)}</select></Field>
        <Field label="Horario"><input value={draft.horario} onChange={(e) => set('horario', e.target.value)} placeholder="Ej. 11:50–12:50 o 07:40–10:50" /></Field>
        <Field label="Tipo de horario"><select value={draft.tipo} onChange={(e) => { const tipo = e.target.value as Block['tipo']; setDraft((d: Block) => ({ ...d, tipo, turno: tipo === 'Taller' ? 'Taller' : 'Teoría', materia: tipo === 'Recreo' ? 'Recreo' : d.materia })) }}><option>Teoría</option><option>Taller</option><option>Recreo</option></select></Field>
        <Field label="Grupo"><select value={draft.grupo ?? ''} onChange={(e) => set('grupo', e.target.value)}><option value="">Curso completo</option><option>Grupo A</option><option>Grupo B</option><option>Ambos</option></select></Field>
        <Field label="Situación de revista"><select value={draft.situacionRevista} onChange={(e) => set('situacionRevista', e.target.value)}><option>Titular</option><option>Provisional</option><option>Titular Interino</option></select></Field>
        <Field label="Suplente" wide><input value={draft.suplente ?? ''} onChange={(e) => set('suplente', e.target.value)} placeholder="Opcional" /></Field>
      </div>
      <div className="flex items-center justify-between border-t bg-slate-50 p-4 rounded-b-2xl">
        <Button variant="ghost" className="text-red-600 hover:text-red-700" onClick={onDelete}>Eliminar bloque</Button>
        <div className="flex gap-2">
          <Button variant="outline" onClick={onClose}>Cancelar</Button>
          <Button className="bg-[#0b5c46] text-white hover:bg-[#074030] shadow-md" onClick={() => onSave(draft)}>Guardar bloque</Button>
        </div>
      </div>
    </div>
  </div> 
}

function Field({ label, children, wide }: { label: string; children: React.ReactNode; wide?: boolean }) { 
  return <label className={`flex flex-col gap-1.5 text-xs font-bold text-slate-600 ${wide ? 'sm:col-span-2' : ''}`}>{label}{children}</label> 
}
