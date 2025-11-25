function gerarTabuada() {
    const numero = document.getElementById('numero').value;
    const resultado = document.getElementById('resultado');
    
    resultado.innerHTML = "";

    if (numero <1 || numero >10) {
        resultado.innerHTML = "Por favor, insira um número entre 1 e 10.";
        return;
    }
    for (let i = 1; i <= 10; i++) {
        const linha = `${numero} x ${i} = ${numero * i}<br>`;
        const p = document.createElement('p');
        p.innerHTML = linha;
        resultado.appendChild(p);
    }
}