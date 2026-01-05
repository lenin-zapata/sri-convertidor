import React from 'react'
import { useCallback, useState } from 'react'
import './App.css' 
import './index.css'
import { FileText, UploadCloud, DownloadCloud } from 'lucide-react'
// Use xlsx-js-style to support styles (fills, borders, number formats)
import XLSX from 'xlsx-js-style'

type ItemRecord = {
  codigoPrincipal: string
  cantidad: number
  descripcion: string
  precioUnitario: number
  descuento: number
  precioTotalSinImpuesto: number
  iva: number
}

type InvoiceRecord = {
  // Emisor
  issuerRazonSocial: string
  issuerRUC: string
  issuerNombreComercial?: string
  issuerDirMatriz?: string
  // Autorización
  numeroAutorizacion?: string
  fechaAutorizacion?: string
  claveAcceso?: string
  // Comprador / factura
  razonSocial: string
  razonSocialComprador?: string
  identificacionComprador?: string
  fechaEmision: string
  secuencial: string
  ruc?: string
  estab?: string
  ptoEmi?: string
  noFactura?: string
  importeTotal: number
  iva: number
  totalDescuento?: number
  moneda?: string
  email?: string
  folio?: string
  archivo?: string
  items: ItemRecord[]
}

export default function App(): React.ReactElement {
  const [records, setRecords] = useState<InvoiceRecord[]>([])
  const [processing, setProcessing] = useState(false)
  const [errors, setErrors] = useState<string[]>([])

  // --- NUEVO ESTADO PARA LA PUBLICIDAD ---
  const [adShown, setAdShown] = useState(false)

  const reset = () => {
    setRecords([])
    setErrors([])
  }

  // --- NUEVA FUNCIÓN PARA EL CLIC DEL BOTÓN ---
  const handleSelectFilesClick = (e: React.MouseEvent) => {
    // Si NO se ha mostrado el anuncio todavía...
    if (!adShown) {
      e.preventDefault()
      e.stopPropagation()
      
      // Abrimos tu enlace en una nueva pestaña
      window.open('https://otieu.com/4/10321796', '_blank')
      
      // Marcamos que ya se mostró para la próxima
      setAdShown(true)
      
      // Importante: No hacemos nada más, así que el selector de archivos NO se abre esta vez
      return
    }

    // Si ya se mostró el anuncio, abrimos el selector de archivos normalmente
    document.getElementById('file-input')?.click()
  }

  const parseInvoiceFromXml = (text: string, filename?: string): InvoiceRecord | null => {
    try {
      const parser = new DOMParser()
      let xml = parser.parseFromString(text, 'text/xml')

      // Extraer número y fecha de autorización si vienen fuera (envoltura <autorizacion>)
      let numeroAutorizacion = ''
      let fechaAutorizacion = ''
      try {
        const autoDoc = parser.parseFromString(text, 'text/xml')
        numeroAutorizacion = autoDoc.getElementsByTagName('numeroAutorizacion')[0]?.textContent?.trim() ?? ''
        fechaAutorizacion = autoDoc.getElementsByTagName('fechaAutorizacion')[0]?.textContent?.trim() ?? ''
      } catch (e) {
        // ignore
      }

      // Si el XML está envuelto en <autorizacion><comprobante><![CDATA[...]]></comprobante></autorizacion>
      // extraemos el contenido del CDATA y lo parseamos de nuevo.
      try {
        const comprobanteNode = xml.getElementsByTagName('comprobante')[0]
        const compText = comprobanteNode?.textContent?.trim() ?? ''
        if (compText && (compText.startsWith('<?xml') || compText.includes('<factura'))) {
          xml = parser.parseFromString(compText, 'text/xml')
        }
      } catch (e) {
        // ignore and continue with original xml
      }

      const infoTrib = xml.getElementsByTagName('infoTributaria')[0]
      const infoFac = xml.getElementsByTagName('infoFactura')[0]

      const getText = (parent: Element | undefined | null, tag: string) => {
        try {
          return parent?.getElementsByTagName(tag)[0]?.textContent?.trim() ?? ''
        } catch {
          return ''
        }
      }

      const razonSocial = getText(infoTrib, 'razonSocial')
      const ruc = getText(infoTrib, 'ruc')
      const secuencial = getText(infoTrib, 'secuencial')
      const estab = getText(infoTrib, 'estab')
      const ptoEmi = getText(infoTrib, 'ptoEmi')
      const noFactura = (estab && ptoEmi && secuencial) ? `${estab}-${ptoEmi}-${secuencial}` : secuencial
      const fechaEmision = getText(infoFac, 'fechaEmision')
      const nombreComercial = getText(infoTrib, 'nombreComercial')
      const dirMatriz = getText(infoTrib, 'dirMatriz')

      // Comprador
      const razonSocialComprador = getText(infoFac, 'razonSocialComprador')
      const identificacionComprador = getText(infoFac, 'identificacionComprador')

      // importeTotal
      const importeTotalRaw = getText(infoFac, 'importeTotal')
      const importeTotal = parseFloat((importeTotalRaw ?? '').replace(/,/g, '')) || 0

      // totalDescuento
      const totalDescuentoRaw = getText(infoFac, 'totalDescuento')
      const totalDescuento = parseFloat((totalDescuentoRaw ?? '').replace(/,/g, '')) || 0

      // moneda
      const moneda = getText(infoFac, 'moneda')

      // IVA: buscar totalImpuesto donde codigo == '2'
      let iva = 0
      try {
        const totalConImpuestos = infoFac?.getElementsByTagName('totalConImpuestos')[0]
        if (totalConImpuestos) {
          const impuestos = totalConImpuestos.getElementsByTagName('totalImpuesto')
          let candidate: number | null = null
          for (let i = 0; i < impuestos.length; i++) {
            const imp = impuestos[i]
            const codigoText = imp.getElementsByTagName('codigo')[0]?.textContent?.trim() ?? ''
            const codigoVal = parseInt(codigoText, 10)
            if (!isNaN(codigoVal) && codigoVal === 2) {
              const valorRaw = imp.getElementsByTagName('valor')[0]?.textContent?.trim() ?? '0'
              const valor = parseFloat(valorRaw.replace(/,/g, '')) || 0
              if (valor > 0) {
                candidate = valor
                break
              }
              if (candidate === null) candidate = valor
            }
          }
          if (candidate !== null) iva = candidate
        }
      } catch (e) {
        // ignore
      }

      // infoAdicional campos (EMAIL, FOLIO, etc)
      let email = ''
      let folio = ''
      try {
        const infoAd = infoFac?.getElementsByTagName('infoAdicional')[0]
        if (infoAd) {
          const campos = infoAd.getElementsByTagName('campoAdicional')
          for (let i = 0; i < campos.length; i++) {
            const c = campos[i]
            const name = c.getAttribute('nombre') ?? ''
            const val = c.textContent?.trim() ?? ''
            if (name.toUpperCase() === 'EMAIL') email = val
            if (name.toUpperCase() === 'FOLIO') folio = val
          }
        }
      } catch (e) {
        // ignore
      }

      // clave de acceso (puede estar en infoTributaria)
      const claveAcceso = getText(infoTrib, 'claveAcceso')

      // Extract detalle items
      const items: ItemRecord[] = []
      try {
        const detallesNode = xml.getElementsByTagName('detalles')[0]
        if (detallesNode) {
          const detalles = detallesNode.getElementsByTagName('detalle')
          for (let d = 0; d < detalles.length; d++) {
            const det = detalles[d]
            const codigoPrincipal = det.getElementsByTagName('codigoPrincipal')[0]?.textContent?.trim() ?? ''
            const descripcion = det.getElementsByTagName('descripcion')[0]?.textContent?.trim() ?? ''
            const cantidad = parseFloat(det.getElementsByTagName('cantidad')[0]?.textContent?.trim() ?? '0') || 0
            const precioUnitario = parseFloat(det.getElementsByTagName('precioUnitario')[0]?.textContent?.trim() ?? '0') || 0
            const descuento = parseFloat(det.getElementsByTagName('descuento')[0]?.textContent?.trim() ?? '0') || 0
            const precioTotalSinImpuesto = parseFloat(det.getElementsByTagName('precioTotalSinImpuesto')[0]?.textContent?.trim() ?? '0') || 0
            // IVA per detalle
            let detalleIva = 0
            try {
              const impuestosNode = det.getElementsByTagName('impuestos')[0]
              if (impuestosNode) {
                const impuestoNodes = impuestosNode.getElementsByTagName('impuesto')
                for (let ii = 0; ii < impuestoNodes.length; ii++) {
                  const imp = impuestoNodes[ii]
                  const codigoImpText = imp.getElementsByTagName('codigo')[0]?.textContent?.trim() ?? ''
                  const codigoImpVal = parseInt(codigoImpText, 10)
                  if (!isNaN(codigoImpVal) && codigoImpVal === 2) {
                    detalleIva = parseFloat(imp.getElementsByTagName('valor')[0]?.textContent?.trim() ?? '0') || 0
                    break
                  }
                }
              }
            } catch (e) {
              // ignore
            }

            items.push({ codigoPrincipal, cantidad, descripcion, precioUnitario, descuento, precioTotalSinImpuesto, iva: detalleIva })
          }
        }
      } catch (e) {
        // ignore
      }

      const rec: InvoiceRecord = {
        issuerRazonSocial: razonSocial,
        issuerRUC: ruc,
        ruc: ruc || undefined,
        issuerNombreComercial: nombreComercial || undefined,
        issuerDirMatriz: dirMatriz || undefined,
        numeroAutorizacion: numeroAutorizacion || undefined,
        fechaAutorizacion: fechaAutorizacion || undefined,
        claveAcceso: claveAcceso || undefined,
        razonSocial: razonSocialComprador || razonSocial || '',
        razonSocialComprador: razonSocialComprador || undefined,
        identificacionComprador: identificacionComprador || undefined,
        fechaEmision,
        secuencial,
        estab: estab || undefined,
        ptoEmi: ptoEmi || undefined,
        noFactura: noFactura || undefined,
        importeTotal,
        iva,
        totalDescuento: totalDescuento || undefined,
        moneda: moneda || undefined,
        email: email || undefined,
        folio: folio || undefined,
        archivo: filename,
        items,
      }

      return rec
    } catch (err) {
      return null
    }
  }

  const handleFiles = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return
    setProcessing(true)
    const out: InvoiceRecord[] = []
    const errs: string[] = []
    for (let i = 0; i < files.length; i++) {
      const f = files[i]
      try {
        const text = await f.text()
        // Puede contener multiples facturas; buscamos por tag <factura>
        const parser = new DOMParser()
        const xmlDoc = parser.parseFromString(text, 'text/xml')
        const facturaNodes = Array.from(xmlDoc.getElementsByTagName('factura'))
        if (facturaNodes.length > 0) {
          facturaNodes.forEach((node) => {
            const serialized = node.outerHTML
            const rec = parseInvoiceFromXml(serialized, f.name)
            if (rec) out.push(rec)
          })
        } else {
          const rec = parseInvoiceFromXml(text, f.name)
          if (rec) out.push(rec)
        }
      } catch (e) {
        errs.push(`${f.name}: ${(e as Error).message}`)
      }
    }

    setRecords((prev) => [...prev, ...out])
    setErrors((prev) => [...prev, ...errs])
    setProcessing(false)
  }, [])

  const onDrop = (ev: React.DragEvent) => {
    ev.preventDefault()
    ev.stopPropagation()
    const dt = ev.dataTransfer
    handleFiles(dt.files)
  }

  const onDragOver = (ev: React.DragEvent) => {
    ev.preventDefault()
    ev.stopPropagation()
  }

  const onFileInputChange: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    handleFiles(e.target.files)
    // reset input so same file can be selected again if needed
    e.currentTarget.value = ''
  }

  const downloadExcel = () => {
    if (records.length === 0) return

    // Build AOAs with one row per detalle (flattened). Header block repeated per row as requested.
    const aoa: (string|number)[][] = []

    // Table header as specified
    const header = [
      'Número de Autorización', // 1
      'Fecha Emisión', // 2
      'No. Factura', // 3 (estab-ptoEmi-secuencial)
      'Razón Social Emisor', // 4
      'RUC Emisor', // 5
      'Cliente', // 6
      'Identificación Cliente', // 7
      'Cod. Principal', // 8
      'Cant.', // 9
      'Descripción', // 10
      'Precio Unitario', // 11
      'Descuento', // 12
      'Precio Total', // 13
      'IVA Impuesto', // 14
      'Importe Total Factura', // 15
    ]

    aoa.push(header)

    records.forEach((r) => {
      const noFactura = r.noFactura ?? (() => {
        // try build estab+ptoEmi+secuencial from stored fields or claveAcceso
        const clave = r.claveAcceso ?? ''
        const sec = r.secuencial ?? ''
        let estab = r.estab ?? ''
        let ptoEmi = r.ptoEmi ?? ''
        try {
          if ((!estab || !ptoEmi) && clave && sec && clave.includes(sec)) {
            const idx = clave.indexOf(sec)
            if (idx >= 6) {
              estab = clave.slice(idx - 6, idx - 3)
              ptoEmi = clave.slice(idx - 3, idx)
            }
          }
        } catch (e) {
          // ignore
        }
        if (estab && ptoEmi && sec) return `${estab}-${ptoEmi}-${sec}`
        return sec
      })()

      r.items.forEach((it) => {
        aoa.push([
          r.numeroAutorizacion ?? '',
          r.fechaEmision ?? '',
          noFactura,
          r.issuerRazonSocial ?? '',
          r.issuerRUC ?? '',
          r.razonSocial ?? '',
          r.identificacionComprador ?? '',

          it.codigoPrincipal ?? '',
          it.cantidad.toFixed(2),
          it.descripcion ?? '',
          it.precioUnitario,
          it.descuento,
          it.precioTotalSinImpuesto,
          it.iva,
          r.importeTotal,
        ])
      })
    })

    const ws = XLSX.utils.aoa_to_sheet(aoa)

    // Styles: header row style
    const headerStyle = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb: '1F4E78' } },
      border: {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'thin', color: { rgb: '000000' } },
        left: { style: 'thin', color: { rgb: '000000' } },
        right: { style: 'thin', color: { rgb: '000000' } },
      },
      alignment: { horizontal: 'center', vertical: 'center' },
    }

    // Apply header style
    const range = XLSX.utils.decode_range(ws['!ref']!)
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_cell({ c: C, r: 0 })
      if (!ws[address]) ws[address] = { t: 's', v: '' }
      ws[address].s = headerStyle
    }

    // Apply borders and number formats for numeric columns
    // Column indexes (0-based): 8 Cant, 10 Precio Unitario, 11 Descuento, 12 Precio Total, 13 IVA, 14 Importe Total Factura
    const currencyFmt = '$ #,##0.00'
    const qtyFmt = '0.00'

    for (let R = 1; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ c: C, r: R })
        if (!ws[addr]) continue
        // ensure a style object exists
        ws[addr].s = ws[addr].s || {}
        // add thin border
        ws[addr].s.border = {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } },
        }

        // apply numeric formats
        if (C === 8) { // Cant.
          ws[addr].z = qtyFmt
        }
        if (C === 10 || C === 11 || C === 12 || C === 13 || C === 14) {
          ws[addr].z = currencyFmt
        }
      }
    }

    // Column widths: Authorization first column 45, Description column index 9 (0-based) width 40, rest 15
    const cols: any[] = []
    const totalCols = aoa[0].length
    for (let c = 0; c < totalCols; c++) {
      if (c === 0) cols.push({ wch: 45 })
      else if (c === 9) cols.push({ wch: 40 })
      else cols.push({ wch: 15 })
    }
    ws['!cols'] = cols

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'RIDE')
    const ts = new Date().toISOString().replace(/[:.]/g, '-')
    XLSX.writeFile(wb, `sri-ride-${ts}.xlsx`)
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <header className="bg-white shadow">
        <div className="max-w-5xl mx-auto px-4 py-4 flex items-center gap-3">
          <FileText className="text-sky-600" />
          <h1 className="text-xl font-semibold text-sky-700">SRI XML a Excel Convertidor</h1>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 py-6">
        <div className="bg-gray-100 p-4 My-4 text-center text-sm text-gray-500">
          {/* Monetag: meta added into document head */}
          {/*Espacio para Publicidad*/}
        </div>

        <section className="bg-white rounded shadow p-6">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h2 className="text-lg font-medium text-sky-700">Sube tus facturas SRI (XML)</h2>
              <p className="text-sm text-gray-500">Arrastra y suelta archivos XML o selecciona desde tu equipo. Procesamiento 100% en el navegador.</p>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={handleSelectFilesClick} // <--- CAMBIO AQUÍ: Usamos la nueva función
                className="inline-flex items-center gap-2 bg-white border border-gray-200 text-slate-700 px-3 py-2 rounded hover:bg-gray-50"
              >
                <UploadCloud className="w-4 h-4" />
                <span className="text-sm">Seleccionar archivos</span>
              </button>
              <button
                onClick={() => { reset() }}
                className="ml-2 inline-flex items-center gap-2 bg-red-50 text-red-700 px-3 py-2 rounded border border-red-100 hover:bg-red-100"
              >Limpiar</button>
            </div>
          </div>

          <div
            onDrop={onDrop}
            onDragOver={onDragOver}
            className="border-2 border-dashed border-slate-200 rounded p-8 text-center bg-sky-50"
            style={{ minHeight: 160 }}
          >
            <p className="text-sky-700 font-medium">Arrastra archivos aquí</p>
            <p className="text-sm text-slate-500">También puedes hacer click en "Seleccionar archivos"</p>
            <input id="file-input" type="file" multiple accept=".xml,text/xml,application/xml" className="hidden" onChange={onFileInputChange} />
          </div>

          <div className="mt-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-gray-600">Registros procesados: <strong>{records.length}</strong></p>
                {processing && <p className="text-sm text-gray-500">Procesando archivos...</p>}
              </div>
              <div>
                <button
                  onClick={downloadExcel}
                  disabled={records.length === 0}
                  className="inline-flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded disabled:opacity-40"
                >
                  <DownloadCloud className="w-4 h-4" />
                  <span>Descargar Excel</span>
                </button>
              </div>
            </div>

            {errors.length > 0 && (
              <div className="mt-3 text-sm text-red-600">
                <strong>Algunos archivos fallaron:</strong>
                <ul className="mt-1 list-disc list-inside">
                  {errors.map((e, i) => <li key={i}>{e}</li>)}
                </ul>
              </div>
            )}

            <div className="mt-6">
              <h3 className="text-sm font-medium text-slate-700 mb-2">Previsualización (primeros 5 registros)</h3>
              <div className="overflow-x-auto bg-white border rounded">
                <table className="min-w-full text-sm">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="px-3 py-2 text-left">Razón Social</th>
                      <th className="px-3 py-2 text-left">RUC</th>
                      <th className="px-3 py-2 text-left">Fecha</th>
                      <th className="px-3 py-2 text-left">Secuencial</th>
                      <th className="px-3 py-2 text-right">Importe Total</th>
                      <th className="px-3 py-2 text-right">IVA</th>
                    </tr>
                  </thead>
                  <tbody>
                    {records.slice(0, 5).map((r, idx) => (
                      <tr key={idx} className={idx % 2 === 0 ? 'bg-white' : 'bg-slate-50'}>
                        <td className="px-3 py-2">{r.razonSocial}</td>
                        <td className="px-3 py-2">{r.ruc}</td>
                        <td className="px-3 py-2">{r.fechaEmision}</td>
                        <td className="px-3 py-2">{r.secuencial}</td>
                        <td className="px-3 py-2 text-right">{r.importeTotal.toFixed(2)}</td>
                        <td className="px-3 py-2 text-right">{r.iva.toFixed(2)}</td>
                      </tr>
                    ))}
                    {records.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-3 py-6 text-center text-sm text-slate-500">No hay registros aún.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        </section>

        <div className="bg-gray-100 p-4 My-4 text-center text-sm text-gray-500">
          {/*Espacio para Publicidad*/}
        </div>
      </main>

      <footer className="text-center text-xs text-gray-400 py-6">
        Built with ❤️ for accountants in Ecuador
      </footer>
    </div>
  )
}