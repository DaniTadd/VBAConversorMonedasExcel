<h1> Conversor de monedas VBA para MS Excel </h1>

<h2>Conversor y graficador de monedas según el tipo de cambio multilateral desarrollado con VBA en MS Excel.</h2>

<h2>Introducción</h2>

<p>
Este proyecto tiene por objetivo mejorar habilidades de desarrollo de aplicaciones en código VBA, uso del paradigma de POO, habilidades de escritura y presentación de proyectos.
También se ensayará la selección de las especificaciones necesarias para un proyecto satisfactorio, el análisis de riesgos de la aplicación y la planificación de pruebas que permitan verificar que la aplicación desarrollada es suficientemente robusta como para ser Publicada.
  
La confección de un "Conversor de monedas" se suele utilizar en la práctica de distintos lenguajes de programación.
En este caso, se trata de un proyecto propuesto en el curso "Excel/VBA for creative problem solving, Part 3." Realizado en la plataforma Coursera <sup>TM</sup>

Se agruparán muy brevemente los conceptos de:
<ul>
  <li>URS (Requerimientos elaborados por el usuario para la búsqueda de proveedores, servicios y /o proyectos (prediseñados o no).</li>
  <li>Especificaciones</li>
  <li>Análisis de Riesgos</li>
  <li>Testeo QA de la apliación</li>
</ul>
</p>

<h2>Objetivo</h2>

<p>
Se desea desarrollar una aplicación en código VBA para utilizar en Excel que realice la conversión de múltiples monedas entre si en forma automática y que permita graficar la evolución del tipo de cambio bilateral durante los 30 días anteriores a una fecha seleccionada por el usuario.
Los datos de "ratio de conversión" de monedas deben ser actualizados diariamente.
</p>

<h2>Requerimientos de Usuario</h2>

<ol>
  <li>Debe incluir al menos la conversión entre: peso argentino, dólar estadounidense, euro, libra esterlina, reales, pesos mexicanos, sol, boliviano, guaraní, renminbi, yen y won.</li>
  <li>El usuario debe poder seleccionar cualquier moneda de partida y deseada. Es decir, no se deben necesitar pasos intermedios.</li>
  <li>Se debe poder graficar un período cualquiera de 30 días del tipo de cambio bilateral entre dos monedas cualesquiera que seleccione el usuario.</li>
  <li>En caso de que no hubiera datos del día en curso, la aplicación debe informar al usuario. No debe mostrarse la ventana de depuración del código al usuario.</li>
  <li>La aplicación debe controlar el ingreso equivocado de caracteres no numéricos por el usuario. No debe mostrarse la ventana de depuración del código al usuario.</li>
  <li>Se debe admitir la creación de varios gráficos de monedas diferentes en diferentes "hojas".</li>
</ol>

<h2>Análisis de Riesgos</h2>

<p>Se determinará la criticidad de las funcionalidades.</p>
<p>Cada una de los atributos de riesgo tendrá un valor según la criticidad:</p>
<ul>
  <li>Severidad: de 1 a 3 se evaluará el nivel de severidad del fallo de la función.</li>
    <ul>
      <li>Siendo 1 una función cuya falla no impide el funcionamiento de la aplicación. No es una función esencial.</li>
      <li>Siendo 2 una función cuya falla impide el funcionamiento de alguna o algunas funcionalidades de la aplicación, pero no de la aplicación central. Es una función importante pero no es esencial.</li>
      <li>Siendo 3 una función cuya falla impide el funcionamiento de la aplicación. Es una función esencial.</li>
    </ul>
  <li>Probabilidad de Ocurrencia: de 1 a 3 se evaluará el nivel de Probabilidad de ocurrencia del fallo de la función.</li>
    <ul>
      <li>Siendo 1 una función cuya ocurrencia depende del código. Si el código se ha testeado y documentado correctamente en la etapa de validación, la funcionalidad tiene bajas o nulas probabilidades de ocurrencia.</li>
      <li>Siendo 2 una función cuya ocurrencia depende de la interacción del usuario con la aplicación, no depende enteramente del código, es "parametrizable". El error de uso aumenta la probabilidad de Ocurrencia, deja de ser una función "automática".</li>
      <li>Siendo 3 una función cuya ocurrencia depende de los parámetros ingresados por el usuario. El resultado de la funcionalidad depende enteramente del usuario.</li>
    </ul>
  <li>Detectabilidad: de 1 a 3 se evaluará el nivel de Probabilidad de ocurrencia del fallo de la función.</li>
    <ul>
      <li>Siendo 1 una función cuyo resultado se basa en una funcionalidad estándar y no se introducen parámentros por el usuario..</li>
      <li>Siendo 2 una función cuyo resultado se basa en funciones estándar, pero el usuario debe introducir parámetros para lograr el resultado deseado.</li>
      <li>Siendo 3 una función prácticamente manual, el sistema se limita a guardar los datos ingresados por el usuario.</li>
    </ul>
</ul>

<table>
  <tr>
    <th>ID</th>
    <th>Descripción</th>
    <th>Severidad</th>
    <th>S</th>
    <th>Ocurrencia</th>
    <th>O</th>
    <th>Detectabilidad</th>
    <th>D</th>
    <th>NR</th>
    <th>P1</th>
    <th>P2</th>
    <th>P3</th>
    <th>P4</th>
    <th>Funcionalidad Crítica</th>
  </tr>
  <tr>
    <td>1</td>
    <td>Conversión entre divisas.</td>
    <td>Si no se puede realizar la conversión correcta, la aplicación no sirve.</td>
    <td>3</td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td>1</td>
    <td>Detectabilidad baja. La conversión es automática, depende del código de la aplicación.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td>2</td>
    <td>Realizar gráfico de 30 días de la paridad entre monedas desde una fecha elegida por el usuario.</td>
    <td>la funcionalidad de graficación no es la función principal.</td>
    <td></td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td></td>
    <td>Detectabilidad baja. El cálculo se realiza en forma automática.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>3</td>
    <td>Selección de fechas pasadas.</td>
    <td>Si no se puede realizar la conversión en fechas pasadas la aplicación no sirve para el fin deseado.</td>
    <td></td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td></td>
    <td>Detectabilidad alta. El resultado es comprobable.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>4</td>
    <td>Conversión entre divisas.</td>
    <td>Si no se puede realizar la conversión correcta, la aplicación no sirve.</td>
    <td></td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td></td>
    <td>Detectabilidad alta. El resultado es comprobable.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>5</td>
    <td>Conversión entre divisas.</td>
    <td>Si no se puede realizar la conversión correcta, la aplicación no sirve.</td>
    <td></td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td></td>
    <td>Detectabilidad alta. El resultado es comprobable.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>6</td>
    <td>Conversión entre divisas.</td>
    <td>Si no se puede realizar la conversión correcta, la aplicación no sirve.</td>
    <td></td>
    <td>Si el código funciona correctamente, la probabilidad de ocurrencia se elimina.</td>
    <td></td>
    <td>Detectabilidad alta. El resultado es comprobable.</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>


  
</table>

<h2>Testeos</h2>

<table>
  <tr>
    <th>ID</th>
    <th>Test</th>
    <th>Descripción</th>
    <th>Resultado esperado</th>
    <th>Resultado obtenido.</th>
  </tr>
  <tr>
    <td>1</td>
    <td>Realizar 3 conversiones con la aplicación.</td>
    <td>1) 15 dólaresestadounidenses a pesos argentinos y 1500 pesos argentinos a dólares estadounidenses, 2) 13 euros a libras esterlinas y 13 libras esterlinas a euros, 3) 25 coronas danesas a pesos colombianos, 25 pesos colombianos a coronas danesas.</td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td>2</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>3</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>4</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>5</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
    <tr>
    <td>6</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>


  
</table>










