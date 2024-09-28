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

        // Mensagem personalizada para Variação 3
        const message = `Olá, ${nomeFantasia || razaoSocial}!\n\nTenho novidades que acredito que vão te ajudar a melhorar ainda mais a performance da sua empresa. Depois de analisar algumas estratégias, desenvolvi um plano personalizado que, quando aplicado, pode aumentar suas vendas significativamente em um curto espaço de tempo.\n\nA proposta é simples, mas muito eficiente, baseada em dados e boas práticas do mercado. Outras empresas que seguiram essas recomendações conseguiram manter um crescimento constante.\n\nQue tal marcarmos uma reunião para discutir como podemos aplicar isso na ${nomeFantasia || razaoSocial}? Será um prazer te ajudar nesse processo!`;

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
