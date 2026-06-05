# 📊 Productivity Analysis Dashboard

Aplicación interactiva construida con **Streamlit** para analizar la productividad de equipos a través de dos modelos distintos según el cliente:

| Modelo | Cliente | Métrica | Lógica |
|---|---|---|---|
| 🟢 **Galderma** | Galderma | Story Points | More is Best |
| 🔵 **AMS** | AMS / Softtek | Horas de Esfuerzo | Less is Best |

---

# 🚀 Cómo ejecutar

### Opción A — Ejecutable (usuarios no técnicos)
Descarga `ProductivityAnalysis.exe`, haz doble clic. El navegador abre la app automáticamente.

### Opción B — Línea de comandos
```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
python -m streamlit run main.py
```
#### 🛠️ Dependencias

```
streamlit
pandas
numpy
plotly
openpyxl
```

### Opción C — Streamlit Cloud
```
https://appapp-evlye6wjacvnkpqqpnpyfe.streamlit.app/
```

---

## 📂 Formato del archivo Excel

La app detecta automáticamente las columnas del Excel. La hoja debe llamarse **`RawData`** (o ser la primera hoja del archivo).

### Modelo AMS — Columnas requeridas

| Columna | Descripción | Requerida |
|---|---|---|
| `Assigned To` | Nombre completo del recurso | ✅ Sí |
| `IS` | Siglas del recurso (ej: `CFPN1`) | No |
| `Group` | Grupo o equipo | ✅ Sí |
| `WBS` | Código del proyecto WBS | ✅ Sí |
| `EndDate` | Fecha de cierre del ticket | ✅ Sí |
| `Effort` | Horas invertidas en el ticket | ✅ Sí |
| `Service Type` | Tipo de servicio (Change / Service) | No |
| `Category` | Categoría del ticket | No |

> El sistema también acepta variantes: `Assignee`, `Resource` en lugar de `Assigned To`; `End Date` en lugar de `EndDate`.

### Modelo Galderma — Columnas requeridas

| Columna | Descripción |
|---|---|
| `Developer` | Nombre del desarrollador (acepta múltiples separados por `/`) |
| `Points` | Story points entregados |
| `Period` | Período del mes (fecha) |
| `Status` | Estado del ticket — solo se procesan `Ready to Deploy` y `Closed` |
| `QA Tester` | Tester asignado (opcional) |

---

## ⚙️ Controles del panel lateral

| Control | Descripción |
|---|---|
| **Select productivity model** | Alterna entre Galderma (Points) y AMS (Effort) |
| **Upload Excel file** | Carga el archivo `.xlsx` con los datos |
| **Analyze by** | Dimensión de análisis: `Assigned To`, `IS`, `Group`, `WBS`, `Service Type`, etc. |
| **Select values** | Filtra qué recursos o grupos mostrar en los gráficos |
| **Analysis mode** | `Individual`: una línea por recurso. `Global`: todos combinados en una sola serie |
| **Charts to show** | Selecciona cuáles gráficas renderizar |

---

## 📐 Lógica de cálculo

El motor de productividad compara el desempeño actual contra una **línea base histórica** mediante ventanas deslizantes de períodos mensuales.

### Parámetros de ventana

```
CURRENT_SIZE  = 3  → Meses más recientes que se evalúan
GAP_SIZE      = 3  → Meses de separación entre ventanas (amortiguador)
BASELINE_SIZE = 3  → Meses históricos usados como referencia
```

**Representación en el tiempo (del más reciente al más antiguo):**
```
[ Mes-1, Mes-2, Mes-3 ] → Ventana ACTUAL   (lo que medimos)
[ Mes-4, Mes-5, Mes-6 ] → GAP              (se ignora)
[ Mes-7, Mes-8, Mes-9 ] → Ventana BASELINE (la referencia)
```

> Si el recurso tiene **menos de 9 períodos** de historia, el baseline se toma de los últimos registros disponibles sin respetar el GAP (modo degradado).

---

### Paso 1 — Agregación mensual

Para cada combinación de `(Período, Dimensión)` se calculan tres valores:

| Campo | Descripción |
|---|---|
| `n` | Cantidad de tickets en ese mes |
| `Sum` | Suma total de la métrica (Effort o Points) |
| `Mean` | Promedio de la métrica por ticket |

**Ejemplo AMS:**

| Period | Assigned To | n | Sum (Effort) | Mean |
|---|---|---|---|---|
| Ene-2025 | Juan Pérez | 12 | 96 h | 8.0 h |
| Feb-2025 | Juan Pérez | 15 | 105 h | 7.0 h |
| Mar-2025 | Juan Pérez | 10 | 90 h | 9.0 h |

---

### Paso 2 — EpU de Baseline (Esfuerzo por Unidad)

Mide cuánto esfuerzo/puntos requería en promedio **cada ticket** durante el período de referencia histórica.

```
EpU_BL = Sum_baseline / n_baseline
```

**Ejemplo:**
```
Baseline (Meses 7-8-9):  Sum = 291 h,  n = 37 tickets
EpU_BL = 291 / 37 = 7.86 h por ticket
```

---

### Paso 3 — Esfuerzo Esperado (Base Equivalent)

¿Cuánto esfuerzo habría consumido el volumen actual **si se mantuviera la eficiencia del baseline?**

```
Base_Equiv = EpU_BL × n_actual
```

**Ejemplo:**
```
Ventana actual: n = 37 tickets
Base_Equiv = 7.86 × 37 = 290.8 h  ← lo que "debería" haber tomado
```

---

### Paso 4 — Índice de Productividad

```
Productividad = ((Effort_actual - Base_Equiv) / Base_Equiv) × σ
```

Donde **σ** depende del modelo:

| Modelo | σ | Razón |
|---|---|---|
| AMS — Less is Best | **-1** | Esfuerzo bajo = bueno → el resultado se invierte para que positivo = mejora |
| Galderma — More is Best | **+1** | Puntos altos = bueno → sin inversión |

**Ejemplo AMS (Less is Best):**
```
Effort_actual = 270 h   → gastó menos de lo esperado ✅
Base_Equiv    = 290.8 h

Productividad = ((270 - 290.8) / 290.8) × (-1)
              = (-0.0715) × (-1)
              = +0.0715  →  +7.15%  ← mejora de productividad
```

**Ejemplo con deterioro:**
```
Effort_actual = 320 h   → gastó más de lo esperado ❌
Base_Equiv    = 290.8 h

Productividad = ((320 - 290.8) / 290.8) × (-1)
              = (0.1003) × (-1)
              = -0.1003  →  -10.03%  ← deterioro
```

---

### Cómo leer los gráficos

| Valor | Significado |
|---|---|
| **> 0%** | Mejora respecto al baseline (consume menos o entrega más) |
| **= 0%** | Igual al baseline |
| **< 0%** | Deterioro respecto al baseline |

La **línea punteada negra** en los gráficos marca el 0% (nivel de baseline).

---

### Modo Individual vs Global

| Modo | Comportamiento |
|---|---|
| **Individual** | Calcula productividad por separado para cada recurso. Permite ver quién mejora y quién no. |
| **Global** | Agrega todos los recursos seleccionados en una sola serie. Refleja el desempeño del equipo completo. |

---

## 📊 Gráficos disponibles

| Gráfico | Descripción |
|---|---|
| **Productivity Over Time** | Índice de productividad en % vs el baseline. Línea cero = referencia |
| **Velocity: Real vs Expected** | Compara el esfuerzo/puntos real contra lo que predice el baseline |
| **Count Over Time** | Cantidad de tickets procesados por período |
| **Mean Over Time** | Promedio de esfuerzo/puntos por ticket por período |

