# MR Office — Español (Latinoamérica)

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

<p align="center"><b>Los editores de código abierto de MR Office (documentos, hojas de cálculo, presentaciones, PDF, escritura a mano y código) en el navegador.</b></p>

<p align="center"><a href="https://mroffice.ai/es-419"><b>▶ Prueba la demo en línea, sin instalar nada</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Descargar MR Office desde Releases</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a></p>

<table>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/docs.png" alt="Documentos" width="100%"></a><br><b>Documentos</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/sheets.png" alt="Hojas de cálculo" width="100%"></a><br><b>Hojas de cálculo</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/slides.png" alt="Presentaciones" width="100%"></a><br><b>Presentaciones</b></td>
</tr>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/ink.png" alt="Escritura a mano" width="100%"></a><br><b>Escritura a mano</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/es-419"><img src="../docs/screenshots/code.png" alt="Código" width="100%"></a><br><b>Código</b></td>
</tr>
</table>

MR Office es un espacio de trabajo para notas, IA y archivos reales. Este repositorio publica el código fuente de sus editores: los mismos editores que usa MR Office, que puedes compilar, ejecutar e integrar por tu cuenta. Son editores de código abierto, no una suite ofimática de código abierto completa: la lectura y escritura de archivos de Microsoft Office se queda en MR Office (ver más abajo).

## Los editores

Lo que hace cada editor en este repositorio:

- **Documentos** — un procesador de textos paginado: estilos, fuentes, listas, tablas, encabezados y pies de página, secciones y columnas, formas y ecuaciones, portadas e índices, comentarios y control de cambios, y tinta en la página. Guarda en su propio formato (.mrdoc); abre HTML, Markdown y texto sin formato.
- **Hojas de cálculo** — celdas y fórmulas, formatos de número, formato condicional, ordenación y filtros, paneles inmovilizados y varias hojas, con la suma y el promedio de la selección. La aritmética entre celdas (=A1*B1+10) se calcula en el navegador al escribir; las fórmulas con funciones (SUMA, SI, BUSCARV…) se conservan tal como se escriben. Guarda en su propio formato (.mrsheet); abre y guarda CSV y TSV.
- **Presentaciones** — diapositivas con ocho temas propios, diseños, vistas de patrón y de diseño, texto, formas, cuadros de texto, tablas, WordArt y ecuaciones, comentarios, secciones, transiciones y animaciones con vista previa, y presentación con diapositivas. Guarda en su propio formato (.mrslides).
- **PDF** — leer, buscar, presentar e imprimir archivos PDF, copiar y leer en voz alta su texto, y marcar el texto seleccionado (resaltado, subrayado, tachado, ondulado) con notas y un panel de comentarios mientras el archivo está abierto; al guardar se devuelve el archivo tal como se abrió.
- **Escritura a mano** — notas manuscritas con bolígrafos sensibles a la presión, resaltadores, borradores, lazo y regla, y tinta convertida en formas. Guarda en su propio formato (.mrink).
- **Código** — un editor de código ligero con resaltado de sintaxis para muchos lenguajes, un esquema y buscar y reemplazar; JavaScript/TypeScript, Python, SQL, Lua y shell se ejecutan directamente en el navegador.

El cálculo de fórmulas con funciones, el dibujo de gráficos y la composición de SmartArt los hace el motor propio de MR Office, en la aplicación de escritorio de MR Office: los editores de código abierto conservan las fórmulas, los gráficos y los SmartArt de un archivo, y muestran los SmartArt tal como se guardaron.

## Ejecutar los editores

**▶ [Prueba la demo en línea, sin instalar nada](https://mroffice.ai/es-419)**

Los seis editores están en `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, que comparten `@mrnote/core`) y se ejecutan juntos en `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` inicia un servidor de desarrollo en la misma dirección y `npm run smoke` comprueba una compilación. El README de cada paquete detalla lo que hace y cómo integrarlo.

## Aplicación de escritorio

[`apps/desktop`](../apps/desktop) contiene **MR Office Editors**: el código fuente de la capa de la aplicación de escritorio de MR Office (Electron), con su propia identidad de aplicación. Con el componente gratuito de MR Office de Releases, es la aplicación de escritorio de MR Office, con archivos de Word, Excel, PowerPoint y PDF, inicio de sesión e IA. Sin el componente, solo ejecuta los editores de código abierto, que abren archivos .mrdoc, .mrsheet, .mrslides y .mrink, además de CSV, HTML, Markdown y PDF. La capa y los editores son código fuente AGPL; el componente está compilado, es gratuito y no es de código abierto. Después de `npm run build`, iníciala con:

```bash
npm run desktop
```

## Para todas las funciones, instala la aplicación de escritorio de MR Office (gratis)

La aplicación de escritorio de MR Office es gratuita, incluye estos editores y el servicio completo, y funciona sin conexión en tu propia computadora. Agrega lo que los editores de código abierto no tienen:

- Abrir y guardar archivos de Word, Excel y PowerPoint (.docx, .xlsx, .pptx)
- Calcular fórmulas con funciones (SUMA, SI, BUSCARV …)
- Gráficos
- Edición de SmartArt
- Guardar en el archivo los cambios hechos en un PDF

**[Descargar MR Office desde Releases →](https://github.com/XinaTian/MROffice/releases)** · [o desde mroffice.ai](https://mroffice.ai)

### Archivos de Word, Excel, PowerPoint y PDF

Abrir y guardar archivos de Word, Excel y PowerPoint (.docx, .xlsx, .pptx), y guardar en el archivo los cambios hechos en un PDF, lo ofrece la aplicación de escritorio de MR Office (gratis), que incluye todos estos editores y funciona sin conexión en tu propia computadora:

## Licencia

Código abierto bajo la GNU Affero General Public License v3.0 (LICENSE), con los términos adicionales de NOTICE.md: conservar la atribución «MR Office» y el enlace a https://mroffice.ai, sin derechos sobre el nombre ni el logotipo de MR Office. Hay una licencia comercial disponible en support@mr-note.com. El software y los datos de terceros figuran en THIRD_PARTY_NOTICES.md. Contribuciones: consulta CONTRIBUTING.md.

## Más

Sitio web y el producto completo: **https://mroffice.ai** · Preguntas y comentarios: support@mr-note.com
