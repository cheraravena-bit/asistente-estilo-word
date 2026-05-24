# Asistente de Estilo para Word

Prototipo de complemento para Microsoft Word que lee el texto del documento y muestra sugerencias en una barra lateral sin modificar las palabras del autor.

## Archivos principales

- `manifest.xml`: manifiesto del complemento de Word.
- `src/taskpane/taskpane.html`: interfaz del panel lateral.
- `src/taskpane/taskpane.js`: lectura del documento y generación de sugerencias.
- `SKILL.md`: reglas de estilo que guiarán el asistente.

## Uso previsto

1. Completar `SKILL.md` con las 5 reglas de estilo definitivas.
2. Instalar dependencias con `npm install`.
3. Iniciar el complemento con `npm start`.

El prototipo pide permiso de solo lectura del documento (`ReadDocument`) y no ejecuta reemplazos automáticos.

## Usar Office en web sin localhost

Para no depender de `https://localhost:3000`, publica estos archivos en un hosting HTTPS:

- `src/taskpane/taskpane.html`
- `src/taskpane/taskpane.css`
- `src/taskpane/taskpane.js`
- una carpeta `assets` con iconos PNG si quieres iconos reales

Luego copia `manifest.public.template.xml`, reemplaza todas las apariciones de `PUBLIC_URL` por la URL publicada y carga ese manifiesto en Office.

Ejemplo:

```text
PUBLIC_URL -> https://tu-sitio.netlify.app
```

Office Add-ins requiere HTTPS. No uses `file://` para el panel lateral.
