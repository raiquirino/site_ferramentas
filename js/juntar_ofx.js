function unirOFX() {
  const input = document.getElementById('ofxFiles');
  const files = input.files;

  if (files.length === 0) {
    alert('Selecione pelo menos um arquivo OFX.');
    return;
  }

  const leitor = [];
  for (let i = 0; i < files.length; i++) {
    leitor.push(files[i].text());
  }

  Promise.all(leitor).then(conteudos => {
    const corpoOFX = conteudos.map((texto, index) => {
      return index === 0 ? texto : texto.replace(/<\?OFX[^>]*\?>/, '');
    }).join('\n');

    const blob = new Blob([corpoOFX], { type: 'application/x-ofx' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'unificado.ofx';
    a.click();

    URL.revokeObjectURL(url);
  });
}