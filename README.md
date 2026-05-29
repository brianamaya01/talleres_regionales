# Plan de Acción de Bioeconomía — Talleres Regionales TNC Colombia

Herramienta web interactiva para facilitar los talleres regionales de validación del **Plan de Acción de Bioeconomía de Colombia**. Permite revisar, ajustar y validar participativamente las acciones propuestas, y presentar los actores del evento.

---

## Estructura del proyecto

```
talleres_regionales/
├── index.html               ← Aplicación web (auto-contenida, sin servidor)
├── update_data.py           ← Script para regenerar datos desde los Excel
├── raw/
│   ├── padb_eje_cafetero.xlsx   ← Plan de Acción Detallado de Bioeconomía
│   ├── actores.xlsx             ← Organizaciones y personas participantes
│   └── img/                     ← Logos de organizaciones
├── .nojekyll                ← Desactiva Jekyll en GitHub Pages
├── .gitignore
└── README.md
```

---

## Actualizar datos desde Excel

Cuando se modifique cualquiera de los dos archivos Excel fuente:

```bash
# Primera vez: instalar dependencias
pip install openpyxl

# Actualizar index.html con los datos más recientes
python update_data.py
```

El script:
1. Lee `raw/padb_eje_cafetero.xlsx` → reemplaza `DEFAULT_DATA` en el HTML
2. Lee `raw/actores.xlsx` (hojas `organizaciones` y `personas`) → reemplaza `ORGS_DATA` y `PERSONAS_DATA`
3. Escribe el `index.html` actualizado

Luego hacer **commit y push** para publicar los cambios en GitHub Pages.

---

## Desplegar en GitHub Pages

### Primera vez

```bash
# 1. Inicializar repositorio
git init
git add .
git commit -m "Deploy inicial — Plan de Acción Bioeconomía Eje Cafetero"

# 2. Crear repo en GitHub y vincular
git remote add origin https://github.com/USUARIO/REPO.git
git branch -M main
git push -u origin main
```

### Activar GitHub Pages

En el repositorio de GitHub:  
**Settings → Pages → Source: Deploy from a branch → Branch: `main` / `/ (root)`**

La URL pública será:  
`https://USUARIO.github.io/REPO/`

### Actualizar después de cambios en Excel

```bash
python update_data.py          # regenera index.html
git add index.html
git commit -m "Actualizar datos PADB y actores"
git push
```

GitHub Pages publica automáticamente en ~1 minuto.

---

## Uso en el taller

| Sección | Descripción |
|---------|-------------|
| **📋 Plan de acción** | Revisar las 12 acciones por dimensión (Gobernanza, Económica, Medio Ambiente). Editar, añadir, eliminar y marcar como revisadas. |
| **👥 Actores del evento** | Vista de organizaciones con logos, keywords y datos de contacto. Mapa interactivo con georeferenciación. Popup con detalle de cada organización y sus participantes. |
| **▶ Presentación automática** | Pantalla completa que cicla las organizaciones cada 60 segundos. Las que tienen coordenadas muestran el mapa en la mitad derecha. Controles: `←` `→` navegar · `Espacio` pausar · `Esc` salir. |
| **🌐 Biowiki** | Enciclopedia de bioeconomía colombiana embebida como panel lateral. |

---

## Tecnologías

- HTML/CSS/JavaScript puro — sin servidor, sin build, sin dependencias npm
- [Leaflet.js](https://leafletjs.com/) — mapas interactivos
- [OpenStreetMap](https://www.openstreetmap.org/) — tiles de mapa
- `localStorage` — persiste ediciones entre recargas sin servidor
- [openpyxl](https://openpyxl.readthedocs.io/) — lectura de Excel (solo para el script local)

> **Nota:** Los mapas y la Biowiki requieren conexión a internet. El resto funciona offline.
