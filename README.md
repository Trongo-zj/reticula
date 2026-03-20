# ReticulaApp

ReticulaApp es un lanzador de aplicaciones liviano para Windows construido con WPF (.NET 8). Presenta una reticula de 30 espacios donde se pueden anclar accesos directos, ejecutables o referencias del menu Inicio mediante arrastrar y soltar.

## Caracteristicas principales
- Admite los formatos mas habituales de Windows (`*.lnk`, `*.exe`, `*.bat`, `*.cmd`, `*.ps1`) y referencias `shell:` arrastrandolos desde el Explorador o el menu Inicio.
- Permite reorganizar aplicaciones mediante "drag and drop" dentro de la reticula, moviendo automaticamente otras fichas para liberar espacio.
- Un doble clic sobre un espacio vacio abre un selector de archivos para elegir la aplicacion a anclar. Un doble clic en el fondo hace lo mismo tomando el primer espacio libre disponible.
- Arrastrar una ficha fuera de la ventana o sobre el fondo elimina la aplicacion asociada.
- Muestra cada ficha solo con el nombre de la aplicacion (fuente de 16 px) para conservar un aspecto limpio y uniforme.
- Las seis columnas se colorean por parejas (azul, verde azulado y purpura) para facilitar la localizacion visual.
- Guarda automaticamente la cuadrilla en `%LOCALAPPDATA%\ReticulaApp\tiles.json` y ofrece botones para exportar/importar la configuracion.

## Requisitos previos
- Windows 10/11.
- .NET 8 SDK.

## Ejecucion local
1. Restaurar dependencias (solo la libreria `System.Drawing.Common` incluida en el proyecto):
   ```powershell
   dotnet restore
   ```
2. Compilar y ejecutar:
   ```powershell
   dotnet run --project ReticulaApp
   ```

## Uso
1. Arrastra accesos directos (`*.lnk`), ejecutables u otras entradas compatibles sobre la reticula. La aplicacion se colocara en el primer espacio libre o en el cuadro sobre el que la sueltes.
2. Para reorganizar, arrastra una ficha a otra posicion. Si la casilla destino ya esta ocupada, la aplicacion desplazada se recolocara en el espacio libre mas cercano.
3. Para quitar una aplicacion, arrastra su ficha fuera de la ventana o sueltala sobre el fondo vacio.
4. Si prefieres seleccionar manualmente, haz doble clic en un cuadro vacio (o sobre el fondo) para abrir un cuadro de dialogo que inicia en el menu Inicio.
5. Usa los botones "Exportar cuadrícula" e "Importar cuadrícula" para respaldar la configuracion antes de actualizar y restaurarla en nuevas versiones.

## Estructura del proyecto
- `ReticulaApp.csproj`: definicion del proyecto WPF dirigido a `net8.0-windows`.
- `App.xaml`, `App.xaml.cs`: configuracion basica de la aplicacion WPF.
- `MainWindow.xaml`: define la interfaz con la reticula de 6 columnas por 5 filas y los estilos de los botones.
- `MainWindow.xaml.cs`: logica principal, gestion de arrastres externos, reordenamiento interno, resolucion de accesos directos e iconos.
- `publish/`: artefactos de publicacion (no se versionan generalmente).

## Notas tecnicas
- Los datos de cada ficha se modelan con `AppTile` (`MainWindow.xaml.cs`) que expone propiedades observables (MVVM ligero). Para transfers temporales se usa el struct `AppTileState`.
- Las integraciones con Windows Shell permiten resolver rutas desde accesos directos y entradas del menu Inicio (`Shell IDList Array`).
- La configuracion se guarda en `%LOCALAPPDATA%\ReticulaApp\tiles.json` y el mismo formato JSON se usa para exportar o importar desde el encabezado.
- Las columnas se agrupan visualmente por pares con colores distintos (azul, verde azulado y purpura) para facilitar la orientacion dentro de la reticula.
- El umbral de arrastre se ha triplicado respecto al minimo del sistema para evitar que un clic normal se interprete como inicio de arrastre.

## Ideas de mejora
1. Permitir diferentes tamanos de cuadrilla o paginas adicionales.
2. Agregar atajos de teclado (lanzar con Enter, eliminar con Supr).
3. Sincronizar la configuracion mediante un servicio en la nube o carpetas compartidas.
