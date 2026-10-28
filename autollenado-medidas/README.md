# Autollenado de Medidas de Contratación

Página web que toma la **base del Link de ingreso** y el **formato MEDIDAS_CONTRATACION (.xlsm)** de nómina, y devuelve el formato lleno, idéntico al original (macros, protección, estilos, listas y fórmulas intactas).

- Cruce por **POSICIÓN**: columna A de la base ↔ columna M de la hoja con `POSICION` en M4 y `APELLIDO` en R4.
- Todo corre en el navegador. Los archivos **no se suben a ningún servidor**.
- Es un sitio estático (HTML + JS). No necesita backend, base de datos ni `npm install`.

## Estructura

```
index.html          página
styles.css          estilos (modo claro y oscuro)
app.js              lógica: lectura, cruce, validaciones y escritura del .xlsm
vendor/             SheetJS 0.18.5 (Apache-2.0) y JSZip 3.10.1 (MIT), incluidas localmente
```

## Probar en tu computadora

Opción rápida: abre `index.html` con doble clic.

Opción recomendada (igual a producción), desde la terminal de VS Code en esta carpeta:

```bash
npx serve .
```

y abre la dirección que aparece (por ejemplo http://localhost:3000).

## Publicar y generar el link (Vercel)

**Con GitHub (recomendado, se actualiza solo con cada cambio):**
1. Sube esta carpeta a un repositorio de GitHub (puede ser privado).
2. En https://vercel.com → **Add New → Project** → importa el repositorio.
3. Framework Preset: **Other**. No pongas comando de build. Output directory: `.` (raíz).
4. **Deploy**. Vercel te da un link tipo `https://autollenado-medidas.vercel.app`.

**Sin GitHub (desde la terminal de VS Code):**
```bash
npx vercel          # primera vez: inicia sesión y acepta las opciones por defecto
npx vercel --prod   # publica la versión final y te da el link
```

Alternativa: https://app.netlify.com/drop → arrastra la carpeta y te da el link al instante.

> El link es público para quien lo tenga, pero la página no guarda ni envía datos: sin los archivos no muestra nada. Si quieren restringirlo, Vercel permite protección con contraseña (Settings → Deployment Protection).

## Qué hace con cada colaborador

| Formato | Dato de la base |
|---|---|
| Q–AO | Tratamiento, apellidos, nombres, nacimiento, estado civil, nacionalidad, cédula, banco (PRODUBANCO), cuenta, dirección, ciudad, teléfono, provincia, estudios, correo, celular |
| AP–AY / BC–BL | Carga 1 y carga 2 (cónyuge primero, luego hijos) |
| AT / BG | Sexo de la carga: se elige en la página (el formulario no lo pide) |

Supuestos: provincia PICHINCHA salvo ciudad de otra provincia; título (AM) según nivel y si la fecha de fin ya pasó; "N/A", "No tengo", "0000" no cuentan como cargas.

## Validaciones (no se cargan y salen en rojo)

- Posición repetida en la base (no se carga ninguna de las dos).
- Posición vacía, con error de fórmula o que no existe en el formato.
- Fila con datos corridos de columna (tipo de cuenta, correo, nivel de estudios, cuenta = cédula, fecha).
- Si el formulario cambia de columnas, la página se detiene con un mensaje en vez de llenar mal.

## Importante

No subas bases ni formatos reales al repositorio: tienen cédulas y cuentas. El `.gitignore` ya bloquea `.xlsx`, `.xlsm`, `.csv` y `.zip`.
