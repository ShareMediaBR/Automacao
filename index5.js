const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const ExcelJS = require('exceljs');

const client = new Client({
    authStrategy: new LocalAuth()
});

let repliedContacts = new Set();

client.on('qr', qr => {
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
    console.log('Client is ready!');
    processExcel();
});

client.on('message', async (message) => {
    const senderNumber = message.from.split('@')[0];
    console.log(`Received message from ${senderNumber}: ${message.body}`);
    repliedContacts.add(senderNumber);
    await removeContact(senderNumber);
});

client.initialize();

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function processExcel() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('leads.xlsx');
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return;
    }
    const worksheet = workbook.getWorksheet(1);

    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const cnpj = row.getCell(1).value;
        const razaoSocial = row.getCell(2).value;
        const nomeFantasia = row.getCell(3).value;
        const email = row.getCell(4).value;
        const ddd = row.getCell(5).value.toString().replace(/[^\d]/g, '');
        let telefone = row.getCell(6).value.toString().replace(/[^\d]/g, '');

        if (telefone.length === 8 || telefone.length === 9) {
            telefone = `${ddd}${telefone}`;
        }

        if (telefone.length === 11) {
            telefone = `55${telefone.slice(0, 2)}${telefone.slice(2)}`;
        } else if (telefone.length === 10) {
            telefone = `55${telefone}`;
        } else {
            console.error('Invalid phone number length:', telefone);
            continue;
        }

        const formattedNumber = `${telefone}@c.us`;

        if (repliedContacts.has(telefone)) {
            console.log(`Skipping ${formattedNumber} as they have already responded.`);
            continue;
        }

        // Mensagem personalizada para Variação 4
        const message = `Oi, ${nomeFantasia || razaoSocial}! \n\nFiz uma análise do mercado e vi que a ${nomeFantasia || razaoSocial} tem um grande potencial de crescimento.\n\nA estratégia que eu criei é completamente customizada para a sua empresa, com ações que vão ajudar a captar mais clientes e aumentar suas vendas rapidamente.\n\nQue tal agendarmos um bate-papo rápido para eu te mostrar como você pode aplicar isso e ver resultados em um curto prazo?\n\nTenho certeza de que essa parceria vai ser muito benéfica para ambos os lados. Fico no aguardo da sua resposta!`;

        try {
            const response = await client.sendMessage(formattedNumber, message);
            console.log('Message sent to', formattedNumber, response);
        } catch (err) {
            console.error('Error sending message to', formattedNumber, err.response ? err.response.data : err.message);
            continue;
        }

        await delay(120000);
    }
}

async function removeContact(number) {
    console.log(`Removing contact: ${number}`);
}
