const express = require('express');
const app = express();

const gerador = require('./gerador');

app.get('/', (req, res) => {
    res.render('index.hbs')
})

app.get('/excel', gerador.gerarXlsx)

app.listen(3000, () => {
    console.log('Servidor rodando na porta 3000');
})
