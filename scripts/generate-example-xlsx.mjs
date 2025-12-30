import fs from 'fs/promises'
import path from 'path'
import XLSX from 'xlsx'

function extractTag(text, tag) {
  // when using RegExp constructor, backslashes must be escaped
  const re = new RegExp(`<${tag}>([\\s\\S]*?)<\\/${tag}>`, 'i')
  const m = text.match(re)
  return m ? m[1].trim() : ''
}

function extractAllTags(text, tag) {
  const re = new RegExp(`<${tag}>([\\s\\S]*?)<\\/${tag}>`, 'ig')
  const out = []
  let m
  while ((m = re.exec(text)) !== null) {
    out.push(m[1].trim())
  }
  return out
}

function extractImpuestos(text) {
  const impuestos = []
  const blocks = text.match(/<totalImpuesto>[\s\S]*?<\/totalImpuesto>/ig) || []
  for (const b of blocks) {
    const codigo = extractTag(b, 'codigo')
    const valor = parseFloat(extractTag(b, 'valor').replace(/,/g, '')) || 0
    impuestos.push({ codigo, valor })
  }
  return impuestos
}

function extractCampoAdicional(inner, name) {
  const re = new RegExp(`<campoAdicional\\s+nombre="${name}"[\\s\\S]*?>([\\s\\S]*?)<\\/campoAdicional>`, 'i')
  const m = inner.match(re)
  return m ? m[1].trim() : ''
}

async function main() {
  const file = path.resolve('./ejemplos/2912202501179071031900121395110000072325658032311.xml')
  const txt = await fs.readFile(file, 'utf8')

  // extract factura XML: try direct <factura> block, otherwise look for CDATA inside <comprobante>
  let inner = txt
  const facturaMatch = txt.match(/<factura[\s\S]*?<\/factura>/i)
  if (facturaMatch) {
    inner = facturaMatch[0]
  } else {
    const m = txt.match(/<comprobante>\s*<!\[CDATA\[([\s\S]*?)\]\]>\s*<\/comprobante>/i)
    if (m) inner = m[1]
  }

  // also extract numeroAutorizacion and fechaAutorizacion from wrapper if present
  const numeroAutorizacion = extractTag(txt, 'numeroAutorizacion')
  const fechaAutorizacion = extractTag(txt, 'fechaAutorizacion')

  // debug: print snippet of inner
  console.log('--- Inner preview ---')
  console.log(inner.slice(0, 400))
  console.log('--- End preview ---')

  // Now try to find tags
  const issuerRazon = extractTag(inner, 'razonSocial')
  const issuerRuc = extractTag(inner, 'ruc')
  const issuerNombreComercial = extractTag(inner, 'nombreComercial')
  const issuerDirMatriz = extractTag(inner, 'dirMatriz')

  console.log('extracted razonSocial:', issuerRazon)
  console.log('direct match on inner for razonSocial:', inner.match(/<razonSocial>([\s\S]*?)<\/razonSocial>/i))
  console.log('extracted ruc:', issuerRuc)
  console.log('extract secuencial direct:', inner.match(/<secuencial>([\s\S]*?)<\/secuencial>/i))
  console.log('extracted secuencial:', extractTag(inner, 'secuencial'))
  console.log('extracted fechaEmision:', extractTag(inner, 'fechaEmision'))

  // quick sanity test for regex function
  const testStr = '<razonSocial>PRUEBA</razonSocial>'
  console.log('regex sanity:', testStr.match(/<razonSocial>([\s\S]*?)<\/razonSocial>/i))

  const sec = extractTag(inner, 'secuencial')
  const fecha = extractTag(inner, 'fechaEmision')
  const importe = parseFloat(extractTag(inner, 'importeTotal').replace(/,/g, '')) || 0
  const razonComprador = extractTag(inner, 'razonSocialComprador')
  const identificacionComprador = extractTag(inner, 'identificacionComprador')

  const impuestos = extractImpuestos(inner)
  let iva = 0
  const impuestos2 = impuestos.filter(i => i.codigo === '2')
  if (impuestos2.length > 0) {
    const pos = impuestos2.find(i => i.valor > 0)
    iva = pos ? pos.valor : impuestos2[0].valor
  }

  const email = extractCampoAdicional(inner, 'EMAIL')
  const folio = extractCampoAdicional(inner, 'FOLIO')

  // Build a more presentable sheet (header block + table) to resemble the PDF layout
  const aoa = []
  aoa.push(['Emitente Razón Social', issuerRazon])
  aoa.push(['Emitente RUC', issuerRuc])
  if (issuerNombreComercial) aoa.push(['Nombre Comercial', issuerNombreComercial])
  if (issuerDirMatriz) aoa.push(['Dirección Matriz', issuerDirMatriz])
  aoa.push(['Número de Autorización', numeroAutorizacion])
  aoa.push(['Clave de Acceso', extractTag(inner, 'claveAcceso')])
  aoa.push(['Fecha Autorización', fechaAutorizacion])
  aoa.push([])

  // Table header (flattened detail rows)
  aoa.push(['Número de Autorización','Fecha Emisión','No. Factura','Razón Social Emisor','RUC Emisor','Cliente','Identificación Cliente','Cod. Principal','Cant.','Descripción','Precio Unitario','Descuento','Precio Total','IVA Impuesto','Importe Total Factura'])

  // data rows (one per detalle)
  // find detalles
  const detallesMatch = inner.match(/<detalles>[\s\S]*?<\/detalles>/i)
  let detallesBlock = ''
  if (detallesMatch) detallesBlock = detallesMatch[0]
  const detalles = detallesBlock.match(/<detalle>[\s\S]*?<\/detalle>/ig) || []
  for (const det of detalles) {
    const codigoPrincipal = extractTag(det, 'codigoPrincipal')
    const descripcion = extractTag(det, 'descripcion')
    const cantidad = parseFloat(extractTag(det, 'cantidad').replace(/,/g, '') || '0') || 0
    const precioUnitario = parseFloat(extractTag(det, 'precioUnitario').replace(/,/g, '') || '0') || 0
    const descuentoDet = parseFloat(extractTag(det, 'descuento').replace(/,/g, '') || '0') || 0
    const precioTotalSinImpuesto = parseFloat(extractTag(det, 'precioTotalSinImpuesto').replace(/,/g, '') || '0') || 0
    // iva per detail
    let ivaDet = 0
    const impuestosMatch = det.match(/<impuesto>[\s\S]*?<\/impuesto>/ig) || []
    for (const imp of impuestosMatch) {
      const cod = extractTag(imp, 'codigo')
      if (cod === '2') {
        ivaDet = parseFloat(extractTag(imp, 'valor').replace(/,/g, '') || '0') || 0
        break
      }
    }

    aoa.push([numeroAutorizacion, fecha, sec, issuerRazon, issuerRuc, razonComprador, identificacionComprador, codigoPrincipal, cantidad.toFixed(2), descripcion, precioUnitario, descuentoDet, precioTotalSinImpuesto, ivaDet, importe])
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa)
  ws['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 30 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 30 }, { wch: 22 }]
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Factura')
  const outPath = path.resolve('./ejemplos/ejemplo-output.xlsx')
  XLSX.writeFile(wb, outPath)
  console.log('Wrote', outPath)
}

main().catch(err => { console.error(err); process.exit(1) })
