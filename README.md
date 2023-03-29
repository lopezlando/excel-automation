1_ instalar nodeJS y NPM (versión LTS): https://nodejs.org/en

Para verificar que se hayan instalado, abrir una consola cmd y tipear los siguientes comandos:

node -v

npm -v

ambos deberían mostrar la versión de cada programa.

2_ abrir una consola de comando en la carpeta principal de este programa y correr el comando: 

npm i

esto instala las librerías necesarias

3_ en esa misma consola en esta carpeta, correr el comando:

node script.js

Este comando corre el programa. Ejecutar cada vez que se quiera correr el script de manera manual.

Para crear un ejecutable del programa, correr los siguientes comandos:

npm i -g pkg  (correr una sola vez para instalar la libreria de manera global)

pkg script.js (correr cada vez que se quiera crear un ejecutable. crea ejecutables para linux, mac y windows. eliminar los que no se deseen)

Cada vez que se hagan cambios al codigo y se quiera renovar el ejecutable, correr este ultimo comando.