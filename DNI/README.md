# Verificador offline de DNI

Aplicación estática para revisar lotes de escaneos de DNI sin subir archivos a Internet.

## Uso

Abrir `index.html` en el navegador, arrastrar archivos PDF o imágenes y pulsar `Comprobar`.

La aplicación genera informes descargables en HTML, CSV y JSON.

## Alcance de la comprobación

La app realiza una verificación técnica offline:

- formato real del archivo;
- hash y duplicados;
- resolución, proporción, contraste, exposición y detalle de imágenes;
- estructura básica de PDF, páginas, cifrado, acciones incrustadas y texto compatible con DNI/MRZ;
- identificación de frente y reverso dentro del mismo PDF multipágina;
- identificación de PDFs de una página con varias imágenes incrustadas;
- identificación aproximada de dos caras dentro de una misma imagen por proporción;
- detección automática por defecto de solo frente, solo reverso o frente y reverso;
- agrupación de frente y reverso por nombre de archivo, análisis visual o selector manual.
- extracción local de texto PDF con PDF.js para validar números aunque estén comprimidos internamente;
- validación estricta de la letra de control del DNI cuando se puede leer el número.
- filtrado de falsos positivos tipo fecha/metadato, por ejemplo `20260502T`.

Un documento completo ya no se marca como correcto solo por tener frente y reverso: también necesita un dato verificable. Si es un PDF escaneado completo pero sin OCR suficiente, queda a revisión, no como falso. Si es una imagen de baja resolución sin número/letra verificable, baja a no válido. Si la letra de control no encaja en contexto real de DNI/NIF/documento, queda como no válido.

No confirma autenticidad oficial. Para eso haría falta validar contra fuentes oficiales, certificados electrónicos o inspección documental especializada.
