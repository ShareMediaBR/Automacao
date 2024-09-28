const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const ExcelJS = require('exceljs');

// Inicializa o cliente WhatsApp
const client = new Client({
    authStrategy: new LocalAuth()
});

let repliedContacts = new Set(); // Armazena os contatos que responderam

client.on('qr', qr => {
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
    console.log('Client is ready!');
    processExcel(); // Inicia o processamento da planilha assim que o cliente estiver pronto
});

// Adiciona um ouvinte para mensagens recebidas
client.on('message', async (message) => {
    const senderNumber = message.from.split('@')[0]; // Obtém o número do remetente
    console.log(`Received message from ${senderNumber}: ${message.body}`);

    // Adiciona o número ao conjunto de contatos que já responderam
    repliedContacts.add(senderNumber);

    // Aqui você pode implementar a lógica para remover o número da lista
    await removeContact(senderNumber);
});

// Inicializa o cliente
client.initialize();

// Função para adicionar delay
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Função para processar a planilha Excel
async function processExcel() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('leads.xlsx'); // Nome do arquivo Excel
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return;
    }
    const worksheet = workbook.getWorksheet(1);

    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const cnpj = row.getCell(1).value; // CNPJ
        const razaoSocial = row.getCell(2).value; // Razão Social
        const nomeFantasia = row.getCell(3).value; // Nome Fantasia
        const email = row.getCell(4).value; // Email
        const ddd = row.getCell(5).value.toString().replace(/[^\d]/g, ''); // DDD
        let telefone = row.getCell(6).value.toString().replace(/[^\d]/g, ''); // Remover todos os caracteres não numéricos

        // Verifica se o número está sem DDD (8 ou 9 dígitos)
        if (telefone.length === 8 || telefone.length === 9) {
            telefone = `${ddd}${telefone}`; // Adiciona o DDD ao número
        }

        // Formata o número para o formato esperado pelo WhatsApp
        if (telefone.length === 11) {
            // Número com DDD e nono dígito (sem duplicar o 9)
            telefone = `55${telefone.slice(0, 2)}${telefone.slice(2)}`; 
        } else if (telefone.length === 10) {
            // Número fixo com DDD (sem nono dígito)
            telefone = `55${telefone}`; 
        } else {
            console.error('Invalid phone number length:', telefone);
            continue;
        }

        const formattedNumber = `${telefone}@c.us`;

        // Verifica se o contato já respondeu
        if (repliedContacts.has(telefone)) {
            console.log(`Skipping ${formattedNumber} as they have already responded.`);
            continue; // Pula o contato
        }

        const message = `Olá, ${nomeFantasia}\nAnalisando as estratégias de comunicação e conquista de novos clientes da ${nomeFantasia}, percebi que ainda não tem uma marca forte na parte digital.\n\nSegundo o IBGE, 40% das empresas locais como a ${nomeFantasia}, encerram as atividades porque não têm uma parte de marketing digital bem estruturada. Você consegue garantir a sobrevivência da sua empresa nos próximos anos?\n\nCom isso, montei uma estratégia simples e assertiva para você conquistar novos clientes, quero te apresentar esta solução gratuita que criei para você.\n\nAtualmente já ajudei mais de 60 empresas como a ${nomeFantasia} aqui no Paraná. Posso citar a Prorelax, Bridge e CrediPronto que são alguns dos nossos cases de sucesso.\n\nVamos conversar para mostrar como posso te ajudar a ter mais clientes?`;

        // Tenta enviar a mensagem
        try {
            const response = await client.sendMessage(formattedNumber, message);
            console.log('Message sent to', formattedNumber, response);
        } catch (err) {
            console.error('Error sending message to', formattedNumber, err.response ? err.response.data : err.message);
            continue; // Ignora e passa para o próximo contato
        }

        // Aguarda 2 minutos antes de enviar a próxima mensagem
        await delay(120000);
    }
}

// Função para remover o contato da lista
async function removeContact(number) {
    console.log(`Removing contact: ${number}`);
    // Lógica para remover da lista pode ser implementada aqui
}
