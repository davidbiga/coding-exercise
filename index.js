const DocumentService = require('./src/services/documentService');

async function main() {
    const documentService = new DocumentService();

    // Example usage
    const contracts = [
        './contracts/contract1.docx',
        './contracts/contract2.docx',
        './contracts/contract3.docx'
    ];

    const clausesToInsert = {
        default: 'Standard legal clause text...',
        liability: 'Liability clause text...',
        confidentiality: 'Confidentiality clause text...'
    };

    try {
        const updatedFiles = await documentService.processContracts(contracts, clausesToInsert);
        console.log('Updated files:', updatedFiles);
    } catch (error) {
        console.error('Error:', error);
    }
}

main();
