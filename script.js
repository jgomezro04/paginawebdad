let rating = 0;

document.addEventListener('DOMContentLoaded', function() {
  loadComments();
});

function rate(stars) {
  rating = stars;
  const starElements = document.querySelectorAll('.star');
  starElements.forEach((star, index) => {
    if (index < stars) {
      star.classList.add('active');
    } else {
      star.classList.remove('active');
    }
  });
}

function addComment() {
  const nameInput = document.getElementById('name-input');
  const commentInput = document.getElementById('comment-input');
  const emojiSelect = document.getElementById('emoji-select');
  
  const name = nameInput.value.trim();
  const commentText = commentInput.value.trim();
  const emoji = emojiSelect.value;

  if (name !== '' && commentText !== '') {
    const commentList = document.getElementById('comments-list');
    const newComment = document.createElement('div');
    const currentDate = new Date();
    const formattedDate = currentDate.toLocaleDateString('es-ES', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
    newComment.innerHTML = `<strong>${name}</strong> (${formattedDate})<br>${commentText} ${emoji}<br>Calificación: ${rating} estrellas`;
    commentList.appendChild(newComment);
    saveComment(name, commentText, emoji, rating, currentDate);
    nameInput.value = '';
    commentInput.value = '';
    // Reset rating
    rating = 0;
    const starElements = document.querySelectorAll('.star');
    starElements.forEach((star) => {
      star.classList.remove('active');
    });
  } else {
    alert('Por favor, ingresa tu nombre y escribe un comentario antes de enviar.');
  }
}

function saveComment(name, comment, emoji, rating, date) {
  let comments = JSON.parse(localStorage.getItem('comments')) || [];
  comments.push({ name: name, comment: comment, emoji: emoji, rating: rating, date: date });
  localStorage.setItem('comments', JSON.stringify(comments));
}

function loadComments() {
  let comments = JSON.parse(localStorage.getItem('comments')) || [];
  const commentList = document.getElementById('comments-list');
  comments.forEach(comment => {
    const newComment = document.createElement('div');
    const formattedDate = new Date(comment.date).toLocaleDateString('es-ES', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
    newComment.innerHTML = `<strong>${comment.name}</strong> (${formattedDate})<br>${comment.comment} ${comment.emoji}<br>Calificación: ${comment.rating} estrellas`;
    commentList.appendChild(newComment);
  });
}

function clearComments() {
  localStorage.removeItem('comments');
  document.getElementById('comments-list').innerHTML = '';
}



function generateExcel(statistics) {
    const wb = XLSX.utils.book_new(); // Crear un nuevo libro de Excel
  
    // Crear una hoja de Excel con los datos de las estadísticas
    const ws = XLSX.utils.json_to_sheet([
      { 'Total de Comentarios': statistics.totalComments },
      { 'Promedio de Calificación': statistics.averageRating },
      { 'Comentarios por Calificación': '' },
      { '1 Estrella': statistics.ratingCounts[0] },
      { '2 Estrellas': statistics.ratingCounts[1] },
      { '3 Estrellas': statistics.ratingCounts[2] },
      { '4 Estrellas': statistics.ratingCounts[3] },
      { '5 Estrellas': statistics.ratingCounts[4] }
    ]);
  
    XLSX.utils.book_append_sheet(wb, ws, 'Estadísticas de Comentarios'); // Añadir la hoja al libro
  
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }); // Convertir el libro a un array de bytes
  
    return wbout;
  }
  function sendEmailWithExcel(fileData) {
    const formData = new FormData();
    formData.append('file', new Blob([fileData], { type: 'application/octet-stream' }), 'estadisticas_comentarios.xlsx');
  
    Email.send({
      Host: 'smtp.yoursmtp.com',
      Username: 'yourusername',
      Password: 'yourpassword',
      To: 'recipient@example.com',
      From: 'jairogomezcuervo@gmail.com,ideasysolucionesco@hotmail.com',
      Subject: 'Estadísticas de Comentarios',
      Body: 'Adjunto encontrarás las estadísticas de comentarios.',
      Attachments: [
        {
          name: 'estadisticas_comentarios.xlsx',
          data: formData
        }
      ]
    }).then(
      message => console.log(message)
    );
  }
  if (commentsCount % 5 === 0) {
    const statistics = generateStatistics();
    const excelData = generateExcel(statistics);
    sendEmailWithExcel(excelData);
  }
    function generateStatistics() {
        let comments = JSON.parse(localStorage.getItem('comments')) || [];
        let totalRating = 0;
        let ratingCounts = [0, 0, 0, 0, 0];
    
        comments.forEach(comment => {
        totalRating += comment.rating;
        ratingCounts[comment.rating - 1]++;
        });
    
        const totalComments = comments.length;
        const averageRating = totalRating / totalComments;
    
        return { totalComments, averageRating, ratingCounts };
    }
  
  