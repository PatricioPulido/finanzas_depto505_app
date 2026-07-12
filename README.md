# Finanzas · Depto 505

App de registro de gastos e ingresos para el Depto 505, en el mismo formato de tu Excel
(Registro diario + Resumen mensual con presupuesto, diferencia y % usado).

## Correr en local

```bash
npm install
npm run dev
```

Abre http://localhost:5173

## Guardado de datos

Los datos se guardan en el `localStorage` del navegador — es decir, quedan guardados
en ese dispositivo/navegador específico, no en la nube. Si quieres acceder a los mismos
datos desde el celular y el computador, o no perderlos si borras el caché del navegador,
en algún momento conviene migrar a una base de datos real (Supabase, Firebase, etc.) —
avísame cuando quieras dar ese paso y te ayudo.

Mientras tanto, usa el botón de descarga (ícono de flecha hacia abajo) para hacer
backups en JSON de vez en cuando, y el botón de Excel para exportar en el mismo
formato de tu planilla original.

## Desplegar en Vercel

1. Sube este proyecto a un repo de GitHub:
   ```bash
   git init
   git add .
   git commit -m "Finanzas Depto 505"
   git branch -M main
   git remote add origin <URL_DE_TU_REPO>
   git push -u origin main
   ```
2. En [vercel.com](https://vercel.com) → "Add New Project" → importa el repo.
3. Vercel detecta Vite automáticamente (build command `vite build`, output `dist`).
   No necesitas tocar nada, solo dale a "Deploy".
