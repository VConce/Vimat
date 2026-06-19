# Despliegue

## Estado de esta entrega

Esta versión sirve para revisión visual y de contenido. Conserva HTML, scripts
y recursos procedentes de Wix, por lo que no debería considerarse todavía el
reemplazo definitivo de producción.

## Para GitHub Pages

1. Publicar el contenido de esta carpeta en el repositorio configurado.
2. Comprobar `index.html` y rutas alias (`temario`, `contacto`,
   `que-proponemos`, etc.) desde GitHub Pages.
3. Cuando se decida cortar Wix, añadir un `CNAME` con `www.vimat.info`.
4. Cambiar DNS solo después de validar la web publicada.

## Pendientes antes de sustituir Wix

- Convertir HTML espejo en HTML/CSS propio sin runtime de Wix.
- Descargar imágenes y recursos críticos a una carpeta local `assets/`.
- Revisar enlaces sociales reales: la web actual contiene algunos enlaces por
  defecto de Wix.
- Crear redirects o alias definitivos para URLs con tilde, especialmente
  `/inscripción` y `/preciosautónomo`.
- Decidir si el formulario de contacto seguirá siendo `mailto`, Google Forms o
  un backend propio.
- Añadir metadatos SEO finales: descripción, Open Graph, favicon y sitemap.
