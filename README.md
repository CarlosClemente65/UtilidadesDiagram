# UtilidadesDiagram v1.0
## Libreria con diversas utilidades que pueden usarse en aplicaciones para Diagram

### Desarrollado por Carlos Clemente (12-2024)

### Control de versiones
 - Version 1.0 - Primera version funcional.

<br>

**Utilidades incluidas:**
 - QuitaRaros(cadena). Permite quitar los acentos a las vocales y los simbolos raros que pueden dar problemas en el basico; devuelve la cadena ajustada
 - ControlFichero(fichero). Se elimina, si existe, el fichero pasado; devuelve 'true' si se ha borrado el fichero
 - GrabarFichero(fichero, texto). Graba en un fichero el texto pasado con la codificacion 'Default' que es la que usa el basico; devuelve 'true' si se ha grabado el fichero
 - DivideCadena(cadena, divisor). Divide la cadena por el caracter pasado en un maximo de 2 partes (por el primer caracter que encuentra; devuelve dos cadenas con el 'atributo' y 'valor'
 - ChequeoFramework(version). Controla si la version de .NET Framework instalada es superior a la pasada (versiones disponibles 4.7, 4.7.1, 4.7.2, 4.8 y 4.8.1)
<br>
