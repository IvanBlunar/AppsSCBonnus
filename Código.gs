function reenviarCorreos() {
  var idHoja = '15NmAgB2GlVbKLI00fukAqYHpUgHfRPBZXo-vj0LH6Qw'; // ID de tu hoja de cálculo
  var hojaCupones = SpreadsheetApp.openById(idHoja).getSheetByName('Cupones'); // Asegúrate de tener una hoja llamada 'Cupones'
  var cupones = hojaCupones.getRange('A:A').getValues().flat().filter(String); // Obtiene los cupones de la columna A que no estén vacíos

  var query = 'from:webmaster@sanpablo.com.co is:unread'; // Busca correos no leídos del remitente específico
  var threads = GmailApp.search(query, 0, 20); // Limitamos a 20 hilos para evitar exceder los límites de velocidad

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var body = message.getPlainBody(); // Usa getPlainBody() para evitar HTML y simplificar la búsqueda
      
      // Verifica si el correo contiene los ISBN específicos
      var isbnEncontrado = isbn_list.some(function(isbn) {
        return body.includes('ISBN: ' + isbn);
      });

      if (isbnEncontrado) {
        var subject = 'Canjea tu cupón para obtener tu ebook ahora mismo'; // Define el asunto del correo
        var cupon = cupones.shift(); // Obtiene el primer cupón de la lista
        hojaCupones.getRange('A1:A' + cupones.length).setValues(cupones.map(function(c) { return [c]; })); // Actualiza la hoja de cálculos con los cupones restantes

        var recipient = extraerEmail(body); // Llama a una función para extraer el email del cuerpo del mensaje

        if (recipient) {
          var htmlBody = crearHtmlCupon(cupon); // Crea el cuerpo HTML del correo electrónico

          try {
            // Intenta enviar el correo electrónico
            GmailApp.sendEmail(recipient, subject, '', {htmlBody: htmlBody, name: 'San Pablo'});
            message.markRead();
            enviarNotificacion(cupon, recipient); // Llama a la función para enviar notificaciones
            Logger.log('Correo reenviado a: ' + recipient + ' con cupón: ' + cupon);
            
            // Espera 5 segundos entre cada envío para evitar exceder los límites de velocidad
            Utilities.sleep(5000);
          } catch (e) {
            // Maneja el error si la dirección de correo electrónico es inválida
            Logger.log('Error al enviar correo a: ' + recipient + ' - ' + e.message);
          }
        }
      }
    }
  }
}

// Función para enviar notificaciones
function enviarNotificacion(cupon, destinatario) {
  var correosNotificacion = [
    'webmaster@sanpablo.com.co',
    'diplomados@sanpablo.com.co'
  ];
  var asuntoNotificacion = 'Notificación de Reenvío de Cupón';
  var cuerpoNotificacion = 'Se ha reenviado el cupón ' + cupon + ' al correo electrónico: ' + destinatario;
  correosNotificacion.forEach(function(correo) {
    GmailApp.sendEmail(correo, asuntoNotificacion, cuerpoNotificacion);
  });
}

// Lista de ISBNs a buscar en los correos
var isbn_list = [
  '310181', '310182', '310185', '310186', '310188', '310189',
  '310190', '310191', '310187', '310183', '310184'
];

// Función para crear el cuerpo HTML del correo electrónico
function crearHtmlCupon(cupon) {
  // Reemplaza 'url_de_tu_imagen' con la URL de la imagen que deseas usar
  var imageUrl = 'https://bucket.mlcdn.com/a/3866/3866669/images/dd34e2ffd01af50d8da3d4e65951eb118854789a.jpeg';
  return `
    <html>
      <body>
        <div style="text-align: center;">
          <img src="${imageUrl}" alt="Imagen Promocional" width="600" style="max-width:100%;">
          <p style="font-size: 20px;">Tu código de cupón es: <strong>${cupon}</strong></p>
          <div style="text-align: center; margin-top: 20px; margin-bottom: 20px;">
            <a href="https://api.whatsapp.com/send?phone=573154935362" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; font-size: 18px; border-radius: 10px;">Más información</a>
          </div>
        </div>
      </body>
    </html>
  `;
}

// Función para extraer la dirección de correo electrónico del cuerpo del mensaje
function extraerEmail(body) {
  var inicio = body.indexOf('Dirección de correo electrónico:') + 'Dirección de correo electrónico:'.length;
  var fin = body.indexOf('\n', inicio);
  var email = body.substring(inicio, fin).trim();
  
  // Verifica si la cadena de correo electrónico no está vacía
  if (email) {
    // Elimina caracteres no deseados al principio y al final del correo electrónico
    email = email.replace(/^\*+|\*+$/g, '').trim();

    // Validación más perisiva de formato de correo electrónico
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (emailRegex.test(email)) {
      return email;
    } else {
      Logger.log('Correo electrónico inválido: ' + email);
      return null; // Retorna null si el correo no es válido
    }
  } else {
    Logger.log('No se encontró una dirección de correo electrónico en el cuerpo del mensaje.');
    return null;
  }
}
