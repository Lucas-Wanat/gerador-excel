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
                valor: 1000.00
            },
            {
                nome_fantasia: "Empresa 2",
                razao_social: "Empresa 2 LTDA",
                cnpj: "987654321",
                endereco: "Rua 2",
                numero: "321",
                bairro: "Bairro 2",
                cidade: "Cidade 2",
                valor: 2000.00
            },
            {
                nome_fantasia: "Empresa 3",
                razao_social: "Empresa 3 LTDA",
                cnpj: "123123123",
                endereco: "Rua 3",
                numero: "456",
                bairro: "Bairro 3",
                cidade: "Cidade 3",
                valor: 3000.00
            },
            {
                nome_fantasia: "Empresa 4",
                razao_social: "Empresa 4 LTDA",
                cnpj: "321321321",
                endereco: "Rua 4",
                numero: "654",
                bairro: "Bairro 4",
                cidade: "Cidade 4",
                valor: 4000.00
            }
        ];

        let nomeValor = Object.keys(rows[0]);

        nomeValor = nomeValor.map((nome) => {
            return nome.replace(/_/g, " ").toUpperCase();
        })

        await XlsxPopulate.fromBlankAsync()
            .then(workbook => {
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
                let totalValor = 0;
                rows.forEach((row) => {
                    let valores = Object.values(row);
                    linha = linha + 1;
                    valores.forEach((valor, i) => {
                        if(linha % 2 == 0) {
                            sheet.row(linha).cell(i + 1).value(valor).style("fill", { type: "solid", color: "e4f3f9" })
                        } else {
                            sheet.row(linha).cell(i + 1).value(valor).style("fill", { type: "solid", color: "ffffff" })
                        }

                        if(typeof valor == "number") {
                            sheet.row(linha).cell(i + 1).style({ horizontalAlignment: "right" });
                            totalValor += valor;

                            if(rows.length == linha - 1) {
                                if(linha % 2 == 0) {
                                    sheet.row(linha + 1).cell(i).value("TOTAL:").style({ bold: true, horizontalAlignment: "left" });
                                    sheet.row(linha + 1).cell(i + 1).value(totalValor).style({ bold: true, horizontalAlignment: "right" });
                                } else {
                                    sheet.row(linha + 1).cell(i).value("TOTAL:").style({ bold: true, horizontalAlignment: "left" });
                                    sheet.row(linha + 1).cell(i + 1).value(totalValor).style({ bold: true, horizontalAlignment: "right" });
                                }
                            }
                        }
                    })
                })

                return workbook.toFileAsync("./output.xlsx");
            });
        
        await res.download('output.xlsx');
    }
}
