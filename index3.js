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

        // Mensagem personalizada para Variação 2
        const message = `Oi, ${nomeFantasia || razaoSocial}! \n\nEstive pensando sobre a situação da ${nomeFantasia || razaoSocial} e como a estratégia que criei pode não só aumentar as vendas, mas realmente transformar o cenário da sua empresa.\n\nEmpresas que aplicaram essas mesmas ações viram um crescimento significativo em novos clientes e estabilidade financeira. Isto é importante para você?\nPor outro lado, deixar de priorizar esse ajuste no marketing digital pode impactar o faturamento a longo prazo, especialmente em um mercado tão competitivo.\n\nA concorrência está cada vez mais digital e, se você não acompanhar, pode perder espaço rapidamente. Você garante a vida da sua empresa para os próximos 5 anos?\n\nVamos marcar uma conversa para entender como podemos implementar isso e evitar que a ${nomeFantasia || razaoSocial} fique para trás?`;

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
