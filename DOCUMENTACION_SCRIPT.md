# Documentación: Sincronización Google Calendar ↔ Excel

Este documento explica cómo funciona la lógica detrás de `Script_Google_Calendar_Corregido.js`. Su objetivo principal es extraer los turnos de Google Calendar (del calendario "Pacientes"), identificar a qué paciente pertenecen cruzándolos con la hoja maestra `"Pacientes"` y volcarlos prolijamente en la hoja `"Turnos"`.

---

## 1. Lógica de Identificación (Cruce de Pacientes)

Cuando el script lee un turno del calendario (ej: *"María Gómez p 1 vez"*), hace lo siguiente para saber quién es:

1. **Normalización:** Convierte todo el texto a minúsculas, quita espacios extra y borra todas las tildes (ej: "María" se vuelve "maria").
2. **Filtro de Ruido:** Elimina temporalmente palabras comunes (`ctrol`, `p`, `v`, `1 vez`, `presencial`, etc.) para que no interfieran en la matemática de comparación de nombres.
3. **Requisito Excluyente (Apellido):** Compara el texto limpio contra tu lista de Excel. Para que lo considere un candidato, **es obligatorio que el apellido coincida** con una precisión mayor al 75% (permitiendo pequeños errores ortográficos).
4. **Validación Final (Nombre + Apellido):** Si el apellido coincidió, calcula la similitud de todo el texto contra el nombre completo. Si la precisión es mayor al **45%**, considera el cruce **Exitoso**.

---

## 2. Asignación del Nombre y Etiquetas (El "Guardado")

Dependiendo de si logró identificar al paciente o no, el script actúa distinto al grabar en Excel:

### A) Cruce Exitoso (Similitud > 45%)
Si detectó quién es, **borra el texto crudo del calendario** y hace lo siguiente:
1. Pone el **Nombre Oficial** que tenés en la hoja Pacientes (bien escrito y prolijo).
2. **Rescate Dinámico de Extras:** Lee el título original que pusiste en el calendario y separa todas las palabras. Las compara contra el nombre oficial. Cualquier palabra que **NO** forme parte del nombre de la persona, la considera una "Etiqueta Extra" y se la suma al final con un guion (ej: `Certificado`, `Receta`, `Trae estudios`).
3. **Embellecimiento:** Si entre esas palabras extra detecta ciertas abreviaturas, las mejora para que queden prolijas:
   - `p` ➔ `Presencial`
   - `v` ➔ `Virtual`
   - `ctrl` / `ctrol` / `ctr` / `ctl` ➔ `Control`
4. Pone en la columna validación: `"Nombre Validado"`.

### B) Cruce Fallido (No lo encontró)
Si la persona es nueva o el nombre estaba tan mal escrito que no superó el 45% de similitud:
1. **No altera nada:** Graba en la hoja de Turnos el nombre exacto y crudo tal cual lo tipeaste en Google Calendar.
2. Pone en la columna validación: `"Revisar / Paciente Nuevo"`.

---

## 3. Reglas Especiales
 
- **Cancelados:** Si el título del evento original contiene la palabra `"cancela"` o `"cancelado"`, el script automáticamente le adosa la etiqueta `" - Cancelado"` al final de todo el texto en el Excel.
- **Detección de Cambios:** Cada vez que ejecutás el script, compara la fecha de "Última Modificación" del evento de Calendar con la del Excel. Solo sobreescribe o actualiza la fila si realmente cambiaste algo en ese turno (ahorrando tiempo y procesamiento).
- **Turnos Huérfanos:** Si el script detecta en tu Excel un turno de 2026 en adelante que ya no existe en el Google Calendar (porque lo borraste definitivamente), le cambiará el estado a `"No está en calendario"`.
- **Exclusión Fija:** Los eventos del calendario que se llamen **exactamente** `"Virtual"` o `"Presencial - Consultorio"` son ignorados por completo y ni siquiera bajan al Excel.

---

## 4. Estructura de Columnas (Hoja "Turnos")

La hoja `Turnos` consta de 11 columnas organizadas de la siguiente manera:

| Col | Encabezado | Origen / Manejo | Descripción |
| :---: | :--- | :--- | :--- |
| **A** | **Fecha** | Calendar (Sincronizado) | Fecha del turno (`DD.MM.YYYY`) |
| **B** | **Hora** | Calendar (Sincronizado) | Hora del turno (`HH:mm`) |
| **C** | **Paciente** | Calendar / Fuzzy Matching | Nombre oficial validado + etiquetas adicionales |
| **D** | **Validacion** | Script Automático | `Nombre Validado` / `Revisar / Paciente Nuevo` / `No está en calendario` |
| **E** | **Importe** | Manual | Monto en pesos de la sesión |
| **F** | **Medio de Pago** | Manual | Canal de cobro (`NX`, `BANCOR`, `Efectivo`, etc.) |
| **G** | **Estado Pago** | Manual | Estado de cobranza: `PAGADO` o `DEBE` |
| **H** | **Estado Turno** | Manual | Estado de la cita (`Atendido`, `Cancelado`, `Ausente`, etc.) |
| **I** | **Comentario** | Manual | Observaciones (ej: meses abonados, saldos, etc.) |
| **J** | **Event ID** | Google Calendar | Identificador único del evento para evitar duplicados |
| **K** | **Ultima Modificación** | Google Calendar | Timestamp de última edición para detectar cambios |

> **Nota de Seguridad:** Las columnas manuales (**E, F, G, H, I**) nunca son sobreescritas ni borradas por el script de sincronización con Google Calendar.
