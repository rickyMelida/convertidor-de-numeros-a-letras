Funciones para hojas de calculo donde la misma puede convertir de una numero de hasta 9 digitos a letras,
esto se realiza de la siguiente manera.
Se agrega la palabra reservada CONVERTIRALETRAS(numero_a_convertir), y entre los parentesis se debe de colocar los numeros el cual se desea convertir a letras.
Se debe de contemplar que los numeros de dos o mas digitos que lleva cero como el primer digito de la expresion va a ser obviada y el valor en si misma quedara como se expresa sin el cero, por lo que se debe de expresar dichos valores sin los ceros ya que la misma no tendria valor alguno.

FORMA DE USO

Ejemplo:
-Dentro de una hoja de calculo(Excel, LibreOffice Calc, etc) colocar una celda cualquiera, la seguiente expresion:
	=CONVERTIRALETRAS(133)
	
Est ultimo dara como resultado:
	Ciento Treinta Y Tres

Tambien se podria llamar a las funciones independientemente, ya que el mismo proyecto fue pensado para poder reutilizar dichas funciones, es decir si se quiere convertir un numero de dos digitos(por ejemplo 26) solo es cuestion de agregar la siguiente funcion:
	=EsNumeroDeDosDigitos(26)

Esto devolveria:
	Veinti Seis

Se debe de considerar que para que la funcionanalidad corra sin ningun incoveniente, es recomendable utilizar la funcion CONVERTIRALETRAS() con el parametro, segun se menciona mas arriba.
