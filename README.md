# envió AENC afectado por el factor de perdidas.
El AENC es la energía que se encuentra reportada en el ente regulador XM ttps://www.xm.com.co/) y esta se aloja como una matriz de consumos.
Haciendo tratamiento de la información lo que hice fue totalizar esa energía por día y a vez realizo un MERGE con tablas temporales y tablas finales
para obtener el dato final de lo que se requiere.

> Esta energía se debe de afectar por un FACTOR DE PERDIDAS que se encuentra alojado en otros Archivos que se llaman TFROC, se realiza
MERGE de cada uno de los archivos diarios y así obtenemos la información consolidada.

Como este trabajo de forma manual o utilizando herramientas ofimáticas requiere de bastante tiempo, se desarrolló esta herramienta para que le llegue 
la información al Usuario Final.

>Al final lo que se comparte un correo electrónico automático a los que requieren recibir la información. Esta se utiliza para calcular indicadores de perdidas
>![image](https://user-images.githubusercontent.com/20642907/201398655-bce7c858-698a-46f2-8427-220b2aa93e3a.png)

