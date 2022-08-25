// Aqui realizamos a atribuição de funções das bibliotecas para uma variável
const fetch = require('node-fetch');
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Countries List');
// As duas variáveis a seguir são para a estilização da planilha
const style = wb.createStyle({
    font: {
        name: "Arial",
        bold: true,
        color: "#808080",
        size: 12,
    },
    numberFormat: '##0,0',
});

const titleStyle = wb.createStyle({
    font: {
        name: "Arial",
        bold: true,
        color: "#4F4F4F",
        size: 14,
    },
    alignment: {
        horizontal: 'center',
        vertical: 'center'
    },
})

// Aqui é a variável a qual ficará salvo o cabeçalho da coluna
const headingColumnNames = [
    "Name",
    "Capital",
    "Area",
    "Currencies"
];

// Realizamos o fetch da API com os dados
fetch('https://restcountries.com/v2/all')
    .then((data) =>
    // Transformamos os dados coletados em .json
        data.json())
        //Transformamos o "arquivo" com os dados para que o programa possa ler
    .then((completedata) => {
        // console.log(completedata.length);

        // Atribuição das variáveis que serão utilizadas globalmente no programa
        let data1 = "";
        let data2 = "";
        let data3 = "";
        let data4 = "";
        let data5 = "";
        let dataName = [];
        let dataCapital = [];
        let dataArea = [];
        let dataCurrencies = [];

        // Aqui vasculhamos os dados e fazemos algo para cada informação que temos
        completedata.forEach((values) => {
            // o RowIndex... serve para apontar em qual linha está a informação (ela é alterada a cada gravação)
            let rowIndexName = 3;
            // o ColumnIndex... serve para apontar em qual coluna deve ficar a informação (ela não é alterada)
            const columnIndexName = 1

            // os "if" serve para verificar se o valor retorna undefined, e caso retorne, ele executará o que é descrito no if, caso não retorne, ele executará as instruções do else
            if (!values.name) {
                return

            } else {
                // Atribuimos o valor name para esta variável que foi vista anteriormente
                data1 = [values.name]
                // Esta variável irá "empurrar" todos os dados coletados para que possamos distribuir pela planilha
                dataName.push({
                    data1
                })
                // Ele irá realizar a ação descrita para cada dado na variável dataName
                dataName.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        // Aqui é a linha responsável por guardar a informação do dataName na planilha
                        ws.cell(rowIndexName, columnIndexName).string(record[columnName])
                        // Estilização da planilha
                            .style(style)
                    });
                    // Nesta linha, a cada vez que for executado o código, ele irá incrementar o valor de RowIndex... para que bote a próxima informação na próxima linha da planilha
                    rowIndexName++;
                });
                // console.log(data1)


            }

            // N

            const columnIndexCapital = 2
            let rowIndexCapital = 3;

            // Nesse caso, alguns dados de capital retornam undefined, portanto, ele irá seguir o que está no "if" para esses retornos e seguirá para o "else" nos casos que não retorne undefined
            if (!values.capital) {

                // Essa linha é responsável por escolher qual resultado será gravado na planilha, porém, como values... está retornando undefined, ele irá escolher o que está do outro lado, que atua para preencher a célula com "-"
                data2 = values.capital || '-'
                dataCapital.push({ data2 })
                const newData2 = dataCapital.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        ws.cell(rowIndexCapital, columnIndexCapital).string(record[columnName])
                            .style(style)
                    })
                    rowIndexCapital++
                })


                return newData2
            } else {
                data2 = [values.capital]
                dataCapital.push({
                    data2
                })
                dataCapital.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        ws.cell(rowIndexCapital, columnIndexCapital).string(record[columnName])
                            .style(style)
                    });
                    rowIndexCapital++;

                });

            }

            const columnIndexArea = 3;
            let rowIndexArea = 3;

            if (!values.area) {
                data3 = values.area || 0
                dataArea.push([data3])

                const newData3 = dataArea.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        ws.cell(rowIndexArea, columnIndexArea).number(record[columnName])
                            .style(style)
                    })
                    rowIndexArea++
                })

                return newData3

            } else {
                data3 = values.area
                dataArea.push({
                    data3
                })

                dataArea.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        ws.cell(rowIndexArea, columnIndexArea).number(record[columnName])
                            .style(style)
                    });
                    rowIndexArea++;
                });

            }

            const columnIndexObject = 4
            let rowIndexObject = 3;

            if (!values.currencies) {
                data4 = values.currencies || '-' 
                console.log(data4)
                dataCurrencies.push({ data4 })
               const newData4 = dataCurrencies.forEach(record => {
                    Object.keys(record).forEach(columnName => {
                        ws.cell(rowIndexCurrencies, columnIndexCurrencies).string(record[columnName])
                            .style(style)
                    })
                    rowIndexCurrencies++
                })

                return newData4

            } else if (!values.currencies[1]) {

                    data4 = values.currencies[0].code

                    dataCurrencies.push({
                        data4
                    })

                    dataCurrencies.forEach(record => {
                        Object.keys(record).forEach(columnName => {
                            ws.cell(rowIndexObject, columnIndexObject).string(record[columnName])
                                .style(style)
                        });
                        rowIndexObject++;
                    });


                } else {
                    data4 = values.currencies[0].code
                    data5 = data4 + ', ' + values.currencies[1].code
                    dataCurrencies.push({ data5 })

                    dataCurrencies.forEach(record => {
                        Object.keys(record).forEach(columnName => {
                            ws.cell(rowIndexObject, columnIndexObject).string(record[columnName])
                                .style(style)
                        });
                        rowIndexObject++;
                    });


                }

        });

        // Essa variável está responsável pela coluna
        let headingColumnIndex = 1;
        // Aqui, ele capta os nomes vistos lá no topo, e para cada dado dentro da variável headingColumnNames, ele grava na célula e pula para a próxima coluna.
        headingColumnNames.forEach(heading => {
            ws.cell(2, headingColumnIndex++).string(heading)
                .style(style)
        });

        // Essa linha é responsável por escrever o título da tabela
        ws.cell(1, 1, 1, 4, true).string('Countries List')
        // Essa linha define a estilização do título, a qual difere da estilização da planilha
            .style(titleStyle)

        // Aqui é a criação do arquivo, a qual irá imprimir tudo que o código manda sob o nome 'tunts.xlsx'
        wb.write('tunts.xlsx')

    })