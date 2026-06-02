# STCLE Analytics Dashboard

Dashboard analítico para el Sindicato de Tripulantes de Cabina STCLE.

El objetivo del proyecto es transformar automáticamente los roles mensuales de tripulantes (Publicado y Ejecutado) en indicadores operacionales, de programación y fatiga, permitiendo a la directiva y a los socios visualizar tendencias, comparaciones y alertas relevantes.

---

# Estructura del Repositorio

```text
stcle-dashboard/
│
├── .github/
│   └── workflows/
│       └── build-data.yml
│
├── data/
│   │
│   ├── raw/
│   │   ├── Publicado_CABLU_MAY26.xlsx
│   │   ├── Roles_EFECT_MAY26_PUB_JUN26.xlsx
│   │   └── ...
│   │
│   ├── kpis.json
│   ├── analytics.json
│   ├── alerts.json
│   └── metadata.json
│
├── scripts/
│   └── build-data.mjs
│
├── assets/
│   ├── logo.png
│   └── icons/
│
├── index.html
├── app.js
├── styles.css
├── package.json
└── README.md
```

---


# Características Principales

## Vista Personal

Permite a cada trabajador comparar su situación con el resto de su cargo.

### Indicadores disponibles

- Horas programadas del mes.
- Horas efectuadas del mes.
- % de utilización.
- Blancos.
- Turnos ASB.
- Turnos HSB.
- Activaciones ASB.
- Activaciones HSB.
- Ratio de activación.
- PSVNC programados.
- PSVNC ejecutados.
- Comparación contra promedio del cargo.
- Evolución histórica personal.
- Ranking relativo dentro del cargo.

---

## Vista Sindicato

Permite a la directiva monitorear tendencias colectivas.

### Indicadores agregados

- Dotación activa.
- Dotación excluida por ausencias.
- Horas promedio CC.
- Horas promedio CCM.
- Blancos promedio.
- Distribución de blancos.
- Distribución de turnos.
- Activación de reservas.
- PSVNC programados.
- PSVNC ejecutados.
- Personas afectadas.
- Evolución histórica mensual.
- Heatmap de actividad.
- Ranking PSVNC.
- Top 10 trabajadores con más turnos.
- Centro de alertas.

---

# Definiciones Operacionales

## Cargos

| Código | Descripción |
|----------|----------|
| CC | Tripulante de Cabina |
| CCM | Jefe de Servicio |

---

## Blancos

Se consideran blancos únicamente:

- B

Se consideran descansos (no blancos):

- BB
- BL
- OFF
- DO

---

## Turnos

### Turno Aeropuerto

Código base:

- ASB

Puede venir acompañado de numeración.

Ejemplos:

- ASB1
- ASB2
- ASB3

Todos se consideran turnos aeropuerto.

---

### Turno Casa

Código base:

- HSB

Puede venir acompañado de numeración.

Ejemplos:

- HSB1
- HSB2
- HSB3

Todos se consideran turnos casa.

---

## Activación de Turnos

Un turno se considera activado cuando:

1. Existe ASB o HSB en el rol Publicado.
2. En el rol Ejecutado aparece cualquier actividad distinta al turno original para ese mismo día.

Ejemplos:

| Publicado | Ejecutado | Activado |
|------------|------------|------------|
| ASB1 | LA123 | Sí |
| HSB2 | CRM | Sí |
| ASB1 | ASB1 | No |
| HSB1 | HSB1 | No |

---

## Ausencias

Se consideran ausencias:

- SICK
- VAC
- OOF

### Regla de exclusión

Todo trabajador que acumule 7 o más días de ausencia en un mismo mes será excluido de:

- Promedios por cargo.
- Rankings.
- Comparaciones.
- Distribuciones estadísticas.

El trabajador seguirá apareciendo en consultas individuales.

---

## Capacitación y Administrativos

Se consideran actividades administrativas o capacitación:

- ADM
- ASCCBT
- CLA
- CRM
- CRS
- DIT
- EMG
- IET
- MCK
- MTC
- MTE
- MTI
- MTU
- SVC

---

# Vuelos

Se consideran vuelos todas las actividades cuyo código comience con:

```text
LA
```

Ejemplos:

```text
LA123
LA456
LA800
```

---

# PSVNC

## Definición

Período de Servicio de Vuelo Nocturno Consecutivo.

Existe PSVNC cuando un vuelo toca total o parcialmente la ventana:

```text
00:30 - 05:30
```

y esto ocurre en dos días calendario consecutivos.

---

## Ejemplos

| Inicio | Término | Cuenta PSVNC |
|----------|----------|----------|
| 22:00 | 01:00 | Sí |
| 04:30 | 08:00 | Sí |
| 23:00 | 00:20 | No |
| 05:31 | 08:00 | No |

---

## Reglas STCLE

### Programado

Máximo permitido:

```text
1 instancia mensual
```

Si se detectan más:

- Generar alerta.

---

### Ejecutado

Máximo permitido:

```text
2 instancias mensuales
```

Si se detectan más:

- Generar alerta.

---

# KPIs Calculados

## Productividad

- Horas Programadas.
- Horas Efectuadas.
- Diferencia absoluta.
- % Utilización.
- Horas promedio por día trabajado.

---

## Blancos

- Blancos por trabajador.
- Promedio blancos por cargo.
- Distribución de blancos.

---

## Turnos

- Turnos ASB.
- Turnos HSB.
- Activaciones ASB.
- Activaciones HSB.
- Ratio activación ASB.
- Ratio activación HSB.

---

## Fatiga

- PSVNC Programados.
- PSVNC Ejecutados.
- Ranking PSVNC.
- Personas afectadas.

---

## Sindicato

- Dotación activa.
- Dotación excluida.
- Evolución histórica.
- Tendencias mensuales.
- Heatmap actividad.
- Ranking turnos.
- Ranking PSVNC.

---

# Arquitectura

## Flujo de Datos

```text
XLSX
 ↓
GitHub Action
 ↓
build-data.mjs
 ↓
JSON Agregados
 ↓
GitHub Pages
 ↓
Dashboard
```
---

# Archivos JSON Generados

## metadata.json

Contiene:

- Fecha de actualización.
- Años disponibles.
- Meses disponibles.
- Cargos disponibles.
- Cantidad de trabajadores.
- Cantidad de registros.

---

## kpis.json

Contiene:

- KPIs por trabajador.
- KPIs por cargo.
- Blancos.
- Turnos.
- Activaciones.
- PSVNC.
- Comparaciones.

---

## analytics.json

Contiene:

- Evolución mensual.
- Heatmaps.
- Rankings.
- Distribuciones.
- Top 10 turnos.
- Top 10 PSVNC.

---

## alerts.json

Contiene:

- PSVNC Programado Excesivo.
- PSVNC Ejecutado Excesivo.
- Ausencias ≥ 7 días.
- Alertas de integridad de datos.

---

# Actualización Mensual

## Paso 1

Subir los nuevos archivos XLSX a:

```text
data/raw/
```

---

## Paso 2

Realizar commit y push.

---

## Paso 3

GitHub Actions ejecutará automáticamente:

```bash
npm run build:data
```

---

## Paso 4

Se regenerarán automáticamente:

```text
data/kpis.json
data/analytics.json
data/alerts.json
data/metadata.json
```

---

## Paso 5

GitHub Pages mostrará los nuevos datos.

No es necesario modificar código.

---

# Tecnologías

- HTML5
- CSS3
- JavaScript ES2024
- Node.js 24
- XLSX
- GitHub Actions
- GitHub Pages

---

# Diseño

## Paleta de colores

Basada en el logo STCLE.

### Primario

Azul corporativo.

### Secundario

Rojo corporativo.

### Apoyo

Grises neutros.

### Estados

- Éxito: Verde.
- Advertencia: Ámbar.
- Alerta: Rojo.

---

# Mantenimiento

## Responsable funcional

Sindicato de Tripulantes de Cabina STCLE.

## Responsable técnico

Proyecto Dashboard STCLE.

---

# Licencia

Uso interno STCLE.
No destinado para distribución pública sin autorización de la organización.
