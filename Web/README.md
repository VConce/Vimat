# Vimat web estática

Primera versión HTML de la web pública de Vimat, generada a partir de
https://www.vimat.info el 19/06/2026.

## Cómo revisar

Abrir `index.html` en un navegador o servir la carpeta con cualquier servidor
estático.

## Qué incluye

- Páginas principales: inicio, inscripción, quiénes somos, nuestro trabajo,
  test, exámenes 2025, temario y contacto.
- Páginas auxiliares detectadas en Wix: precios, recursos, aprobados y `nt`.
- Alias sin extensión para mantener enlaces internos de Wix como `/temario`,
  `/contacto` o `/que-proponemos`.
- Recursos e imágenes enlazados desde Wix/Wixstatic para conservar el aspecto
  visual original.

## Estado actual

Esta entrega es un espejo HTML funcional de Wix. Es útil para ver la web fuera
del editor de Wix y para comparar contenido, pero todavía no es la versión
final independiente: conserva runtime, scripts y recursos externos de Wix.

## Siguiente iteración recomendada

Convertir el espejo a HTML/CSS propio, descargar assets a `assets/`, limpiar
scripts de Wix y dejar rutas compatibles para GitHub Pages.
