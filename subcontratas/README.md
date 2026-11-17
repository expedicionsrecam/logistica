# Recam Subcontratas

Seguimiento y reclamación a subcontratistas a partir del volcado diario del RP (Sheet «Master subcontratas»).

- `index.html` – la página (un solo archivo). URL: https://expedicionsrecam.github.io/logistica/subcontratas/
- `Code.gs` – Apps Script que lee el Sheet y guarda contactos/reclamaciones en él (instrucciones al inicio del archivo).

Uso diario: pegar el volcado del RP en la primera hoja del Sheet; la página lo lee al abrirla, con *Actualizar* y cada 5 minutos.
Los contactos y promesas quedan en la hoja «Seguimiento»; los emails de proveedores, en «Contactos proveedores».
