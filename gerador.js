const XlsxPopulate = require('xlsx-populate');

module.exports = {
    async gerarXlsx(req, res) {
        let rows = [
            {
                nome_fantasia: "Empresa 1",
                razao_social: "Empresa 1 LTDA",
                cnpj: "123456789",
                endereco: "Rua 1",
                numero: "123",
                bairro: "Bairro 1",
                cidade: "Cidade 1",
                dinheiro: 123.45,
                quantidade: 350,
            },
            {
                nome_fantasia: "Empresa 2",
                razao_social: "Empresa 2 LTDA",
                cnpj: "987654321",
                endereco: "Rua 2",
                numero: "321",
                bairro: "Bairro 2",
                cidade: "Cidade 2",
                dinheiro: 123.45,
                quantidade: 90,
            },
            {
                nome_fantasia: "Empresa 3",
                razao_social: "Empresa 3 LTDA",
                cnpj: "123123123",
                endereco: "Rua 3",
                numero: "456",
                bairro: "Bairro 3",
                cidade: "Cidade 3",
                dinheiro: 123.45,
                quantidade: 32,
            },
            {
                nome_fantasia: "Empresa 4",
                razao_social: "Empresa 4 LTDA",
                cnpj: "321321321",
                endereco: "Rua 4",
                numero: "654",
                bairro: "Bairro 4",
                cidade: "Cidade 4",
                dinheiro: 123.45,
                quantidade: 12,
            },
        ];

        let nomeValor = Object.keys(rows[0]);

        nomeValor = nomeValor.map((nome) => {
            return nome.replace(/_/g, " ").toUpperCase();
        })

        totais = {};

        rows.forEach((row) => {
            Object.keys(row).forEach((key) => {
                if(typeof row[key] == "number") {
                    if(totais[key]) {
                        totais[key] = totais[key] + row[key];
                    } else {
                        totais[key] = row[key];
                    }
                }
            })
        });

        await XlsxPopulate.fromBlankAsync()
            .then(async (workbook) => {
                const sheet = workbook.sheet("Sheet1");

                nomeValor.forEach((nome, index) => {
                    sheet.row(1).cell(index + 1).value(nome).style("fill", { type: "solid", color: "76daf3" }).style({ bold: true, horizontalAlignment: "center" });
                    
                    let larguraMax = 0;
                    let larguras = [];

                    rows.map((row) => {
                        larguras.push(row[nome.replace(/ /g, "_").toLowerCase()].length);
                    })

                    larguraMax = larguras.reduce((prev, current) => { 
                        return prev > current ? prev : current; 
                    });

                    larguraMax = larguraMax > nome.length ? larguraMax : nome.length;

                    sheet.column(index + 1).width(larguraMax + 5);
                })

                let linha = 1;
                
                rows.forEach((row) => {
                    let valores = Object.values(row);
                    linha = linha + 1;
                    valores.forEach((valor, i) => {
                        sheet.row(linha).cell(i + 1).value(valor)

                        if(linha % 2 == 0) {
                            sheet.row(linha).cell(i + 1).style("fill", { type: "solid", color: "ffffff" });
                            sheet.row(linha + 1).cell(i + 1).style("fill", { type: "solid", color: "e4f3f9" });
                        } else {
                            sheet.row(linha).cell(i + 1).style("fill", { type: "solid", color: "e4f3f9" });
                            sheet.row(linha + 1).cell(i + 1).style("fill", { type: "solid", color: "ffffff" });
                        }

                        if(typeof valor == "number") {
                            if(rows.length == linha - 1) {
                                sheet.row(linha + 1).cell(1).value("Total").style({ bold: true, horizontalAlignment: "left" });

                                if(Object.keys(totais).includes(Object.keys(row)[i])) {
                                    sheet.row(linha + 1).cell(i + 1).value(totais[Object.keys(row)[i]]).style({ bold: true, horizontalAlignment: "right" });
                                }
                            }
                        }
                    })
                })

                return await workbook.outputAsync();
            })
            .then((data) => {
                let fileName = `output.xlsx`;

                res.attachment(fileName);

                res.send(data);
            })
            .catch((err) => {
                console.log(err);
            });
    }
}
