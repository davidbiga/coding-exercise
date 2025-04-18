const { Document, Paragraph, TextRun, Packer } = require('docx');
const fs = require('fs-extra');
const path = require('path');
const mammoth = require('mammoth');
const officegen = require('officegen');
const { JSDOM } = require('jsdom');

class DocumentService {
    async processContracts(contracts, clausesToInsert) {
        try {
            // Ensure the amended folder exists
            const amendedFolder = path.join(process.cwd(), 'amended');
            await fs.ensureDir(amendedFolder);
            
            const results = [];
            for (const contract of contracts) {
                const updatedDoc = await this.processDocument(contract, clausesToInsert);
                results.push(updatedDoc);
            }
            return results;
        } catch (error) {
            console.error('Error processing contracts:', error);
            throw error;
        }
    }

    async processDocument(documentPath, clausesToInsert) {
        try {
            const outputPath = path.join(process.cwd(), 'amended', path.basename(documentPath));
            const defaultStyle = {
                font_face: 'Times New Roman',
                font_size: 12,
                align: 'left'
            };

            if (documentPath.includes('contract1.docx')) {
                await this.processContract1(documentPath, outputPath, defaultStyle);
            } else if (documentPath.includes('contract2.docx')) {
                await this.processContract2(documentPath, outputPath, defaultStyle);
            } else if (documentPath.includes('contract3.docx')) {
                await this.processContract3(documentPath, outputPath, defaultStyle);
            }
            
            return outputPath;
        } catch (error) {
            console.error(`Error processing document ${documentPath}:`, error);
            throw error;
        }
    }

    async processContract1(inputPath, outputPath, defaultStyle) {
        try {
            // Read the document content
            const result = await mammoth.extractRawText({ path: inputPath });
            const content = result.value;

            // Create a new Word document
            const docx = officegen('docx');
            
            // Set document properties and default styling
            docx.creator = 'Document Service';
            docx.title = 'Contract';
            
            // Configure default paragraph properties
            const defaultStyle = {
                font_face: 'Times New Roman',
                font_size: 12,
                align: 'left'
            };

            const definitionsPattern = /(\bDefinitions\.\s*\n)/i;
            const parts = content.split(definitionsPattern);

            if (parts.length < 2) {
                throw new Error('Definitions section not found in Contract 1');
            }

            // Add content before "Definitions." with proper styling
            const pObj = docx.createP();
            pObj.addText(parts[0].trim(), {
                ...defaultStyle,
                spacing: { before: 240, after: 120 }
            });

            // Add "Definitions." heading with enhanced styling
            const headingP = docx.createP();
            headingP.addText('Definitions.', {
                ...defaultStyle,
                bold: true,
                underline: true
            });

            // Add Affiliate definition with enhanced styling
            const affiliateP = docx.createP();
            affiliateP.addText('A.', {
                ...defaultStyle,
                bold: true
            });
            affiliateP.addText('\t"Affiliate"', {
                ...defaultStyle,
                bold: true
            });
            affiliateP.addText(' means any entity that directly or indirectly controls, is controlled by, or is under common control with a party, where "control" means the possession, directly or indirectly, of the power to direct or cause the direction of the management and policies of such entity, whether through ownership of voting securities, by contract, or otherwise.', {
                ...defaultStyle
            });

            // Add remaining definitions with consistent styling
            if (parts.length > 2) {
                const remainingContent = parts[2];
                const definitions = remainingContent.split(/(?=\s*"[A-Z][^"]+"\s+means)/);
                definitions.forEach((text, index) => {
                    if (text.trim()) {
                        const p = docx.createP();
                        const letter = String.fromCharCode('B'.charCodeAt(0) + index);
                        p.addText(letter + '.', {
                            ...defaultStyle,
                            bold: true
                        });
                        // Split the definition into term and description
                        const matches = text.match(/^(\s*"[^"]+")\s+(means.*)/);
                        if (matches) {
                            p.addText('\t' + matches[1], {
                                ...defaultStyle,
                                bold: true
                            });
                            p.addText(' ' + matches[2], {
                                ...defaultStyle
                            });
                        } else {
                            p.addText('\t' + text.trim(), {
                                ...defaultStyle
                            });
                        }
                    }
                });
            }

            // Generate the document
            return new Promise((resolve, reject) => {
                const out = fs.createWriteStream(outputPath);
                docx.generate(out, {
                    'finalize': function(written) {
                        resolve();
                    },
                    'error': function(err) {
                        reject(err);
                    }
                });
            });
        } catch (error) {
            console.error('Error in processContract1:', error);
            throw error;
        }
    }

    async processContract2(inputPath, outputPath, defaultStyle) {
        try {
            // Read the document content
            const result = await mammoth.extractRawText({ path: inputPath });
            let content = result.value;

            // Thoroughly normalize the content and search text
            const normalizeText = (text) => {
                return text
                    .replace(/[\n\r\t]+/g, ' ')  // Replace newlines, tabs with spaces
                    .replace(/\s+/g, ' ')        // Collapse multiple spaces
                    .replace(/[""]/g, '"')       // Normalize quotes
                    .toLowerCase()               // Convert to lowercase for case-insensitive matching
                    .trim();
            };

            // Look for a paragraph containing these key phrases
            const searchPhrases = [
                "confidential information",
                "as is",
                "receiving party"
            ];

            const normalizedContent = normalizeText(content);
            
            // Split content into paragraphs
            const paragraphs = content.split(/\n\n+/);
            let targetParagraphIndex = -1;
            let targetParagraph = '';

            // Find the paragraph that contains all search phrases
            for (let i = 0; i < paragraphs.length; i++) {
                const normalizedParagraph = normalizeText(paragraphs[i]);
                if (searchPhrases.every(phrase => normalizedParagraph.includes(phrase.toLowerCase()))) {
                    targetParagraphIndex = i;
                    targetParagraph = paragraphs[i];
                    break;
                }
            }

            if (targetParagraphIndex === -1) {
                console.log('Content excerpt:', normalizedContent.substring(0, 200));
                throw new Error('Could not find the confidentiality paragraph in Contract 2.');
            }

            // Split content at the target paragraph
            const beforeParagraphs = paragraphs.slice(0, targetParagraphIndex).join('\n\n');
            const afterParagraphs = paragraphs.slice(targetParagraphIndex + 1).join('\n\n');

            // Create the replacement text in uppercase
            const replaceText = `THE DISCLOSING PARTY IS PROVIDING CONFIDENTIAL INFORMATION ON AN "AS IS" BASIS FOR USE BY THE RECEIVING PARTY AT ITS OWN RISK. THE DISCLOSING PARTY MAKES NO REPRESENTATIONS OR WARRANTIES REGARDING THE ACCURACY OR COMPLETENESS OF THE CONFIDENTIAL INFORMATION. THE DISCLOSING PARTY DISCLAIMS ALL WARRANTIES, WHETHER EXPRESS, IMPLIED OR STATUTORY, INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF TITLE, NON-INFRINGEMENT OF THIRD PARTY RIGHTS, MERCHANTABILITY, OR FITNESS FOR A PARTICULAR PURPOSE.`;

            // Create a new Word document
            const docx = officegen('docx');
            docx.creator = 'Document Service';
            docx.title = 'Contract';

            // Add all content
            const p = docx.createP();
            p.addText(beforeParagraphs + '\n\n' + replaceText + '\n\n' + afterParagraphs, defaultStyle);

            // Generate the document
            return new Promise((resolve, reject) => {
                const out = fs.createWriteStream(outputPath);
                docx.generate(out, {
                    'finalize': function(written) {
                        resolve();
                    },
                    'error': function(err) {
                        reject(err);
                    }
                });
            });
        } catch (error) {
            console.error('Error in processContract2:', error);
            throw error;
        }
    }

    async processContract3(inputPath, outputPath, defaultStyle) {
        try {
            // Read the document content
            const result = await mammoth.extractRawText({ path: inputPath });
            const content = result.value;

            // Debug: Log content length and first 200 characters
            console.log('Document content length:', content.length);
            console.log('First 200 characters:', content.substring(0, 200));

            // Find the end of Section 10
            const section10EndText = "in furtherance of the Business Purpose.";

            // Normalize both content and search text
            const normalizeText = (text) => {
                return text
                    .replace(/[\n\r\t]+/g, ' ')
                    .replace(/\s+/g, ' ')
                    .replace(/[""]/g, '"')
                    .trim();
            };

            const normalizedContent = normalizeText(content);
            const normalizedSearchText = normalizeText(section10EndText);
            console.log('Normalized search text:', normalizedSearchText);
            // Find the position of the text
            const searchIndex = normalizedContent.indexOf(normalizedSearchText);
            if (searchIndex === -1) {
                console.log('Content excerpt:', normalizedContent.substring(0, 200));
                throw new Error('Could not find Section 10 end text');
            }

            // Find the end of the paragraph (next double newline)
            const endOfParagraph = content.indexOf('\n\n', searchIndex + normalizedSearchText.length);
            const insertPosition = endOfParagraph !== -1 ? endOfParagraph : searchIndex + normalizedSearchText.length;

            // Split content at the paragraph boundary
            const firstPart = content.substring(0, insertPosition);
            const secondPart = content.substring(insertPosition);

            // Debug: Log the split points
            console.log('End of first part:', firstPart.slice(-100));
            console.log('Start of second part:', secondPart.slice(0, 100));

            // Create a new Word document
            const docx = officegen('docx');
            docx.creator = 'Document Service';
            docx.title = 'Contract';

            // Add content up to and including Section 10
            const firstPartP = docx.createP();
            firstPartP.addText(firstPart + '\n\n', defaultStyle);

            // Add new Section 11 (Residuals)
            const residualsHeading = docx.createP();
            residualsHeading.addText('11. Residuals', {
                ...defaultStyle,
                bold: true,
                underline: true
            });

            const residualsContent = docx.createP();
            residualsContent.addText('Nothing in this Agreement shall be construed to limit the Receiving Party\'s right to independently develop or acquire products or services without use of the Disclosing Party\'s Confidential Information, nor shall it restrict the use of any general knowledge, skills, or experience retained in unaided memory by personnel of the Receiving Party.\n\n', defaultStyle);

            // Add the remaining content, replacing section numbers
            const finalSection = docx.createP();
            finalSection.addText(
                secondPart
                    .replace(/12\./g, '13.')  // Replace 12. with 13. first
                    .replace(/11\./g, '12.'),  // Then replace 11. with 12.
                defaultStyle
            );

            // Generate the document
            return new Promise((resolve, reject) => {
                const out = fs.createWriteStream(outputPath);
                docx.generate(out, {
                    'finalize': function(written) {
                        resolve();
                    },
                    'error': function(err) {
                        reject(err);
                    }
                });
            });
        } catch (error) {
            console.error('Error in processContract3:', error);
            throw error;
        }
    }

    findInsertionPoints(doc) {
        // This is a simplified version. You'll need to implement the actual logic
        // to find insertion points based on your specific requirements
        const insertionPoints = [];

        // Example: Find sections marked with specific text or formatting
        doc.sections.forEach((section, sectionIndex) => {
            section.children.forEach((paragraph, paragraphIndex) => {
                if (this.isInsertionMarker(paragraph)) {
                    insertionPoints.push({
                        sectionIndex,
                        paragraphIndex,
                        clauseType: this.determineClauseType(paragraph)
                    });
                }
            });
        });

        return insertionPoints;
    }

    isInsertionMarker(paragraph) {
        // Implement logic to identify insertion markers
        // This could be based on specific text, formatting, or other criteria
        return false;
    }

    determineClauseType(paragraph) {
        // Implement logic to determine which type of clause should be inserted
        return 'default';
    }

    extractDefaultStyling(doc) {
        // Find the first regular paragraph to extract default styling
        for (const section of doc.sections) {
            for (const paragraph of section.children) {
                if (paragraph.children?.[0]) {
                    const run = paragraph.children[0];
                    return {
                        font: run.font || 'Times New Roman',
                        size: run.size || 24,
                        spacing: paragraph.spacing || { before: 240, after: 240, line: 360 },
                        alignment: paragraph.alignment || 'left'
                    };
                }
            }
        }
        // Return default styling if none found
        return {
            font: 'Times New Roman',
            size: 24,
            spacing: { before: 240, after: 240, line: 360 },
            alignment: 'left'
        };
    }

    getContextStyle(doc, insertionPoint) {
        const section = doc.sections[insertionPoint.sectionIndex];
        const prevParagraph = section.children[insertionPoint.paragraphIndex - 1];
        const nextParagraph = section.children[insertionPoint.paragraphIndex + 1];
        
        // Prefer previous paragraph's style, fall back to next, or use default
        const sourceParagraph = prevParagraph || nextParagraph;
        if (sourceParagraph?.children?.[0]) {
            const run = sourceParagraph.children[0];
            return {
                font: run.font,
                size: run.size,
                spacing: sourceParagraph.spacing,
                alignment: sourceParagraph.alignment,
                bold: run.bold,
                italic: run.italic,
                underline: run.underline
            };
        }
        return this.extractDefaultStyling(doc);
    }

    async insertClause(doc, insertionPoint, clause, style) {
        const newParagraph = new Paragraph({
            children: [
                new TextRun({
                    text: clause,
                    font: style.font,
                    size: style.size,
                    bold: style.bold,
                    italic: style.italic,
                    underline: style.underline
                })
            ],
            spacing: style.spacing,
            alignment: style.alignment
        });

        doc.sections[insertionPoint.sectionIndex].children.splice(
            insertionPoint.paragraphIndex,
            0,
            newParagraph
        );
    }

    async insertConfidentialityDisclaimer(doc) {
        try {
            if (!doc.sections || !doc.sections[0] || !doc.sections[0].children) {
                throw new Error('Invalid document structure');
            }

            const section = doc.sections[0];  // Get the first section
            for (let i = 0; i < section.children.length; i++) {
                const paragraph = section.children[i];
                const text = this.getParagraphText(paragraph);
                
                if (text.startsWith('11. Confidentiality')) {
                    // Get the style from the existing paragraph
                    const style = paragraph.children[0] ? {
                        font: paragraph.children[0].font || 'Times New Roman',
                        size: paragraph.children[0].size || 24,
                        spacing: paragraph.spacing || { before: 240, after: 240, line: 360 },
                        bold: paragraph.children[0].bold
                    } : this.extractDefaultStyling(doc);

                    // Split the existing text at the first sentence
                    const sentences = text.split(/(?<=\.)\s+/);
                    if (sentences.length >= 1) {
                        const newParagraph = new Paragraph({
                            children: [
                                new TextRun({
                                    text: sentences[0] + ' ',
                                    ...style
                                }),
                                new TextRun({
                                    text: 'The Disclosing Party makes no representations or warranties regarding the accuracy or completeness of the Confidential Information. ',
                                    ...style
                                }),
                                new TextRun({
                                    text: sentences.slice(1).join(' '),
                                    ...style
                                })
                            ],
                            spacing: style.spacing
                        });

                        // Replace the existing paragraph with the new one
                        section.children[i] = newParagraph;
                        return;
                    }
                }
            }

            throw new Error('Section 11 (Confidentiality) not found in the document');
        } catch (error) {
            console.error('Error in insertConfidentialityDisclaimer:', error);
            throw error;
        }
    }

    async insertResidualsClause(doc) {
        try {
            const style = this.extractDefaultStyling(doc);
            
            const heading = new Paragraph({
                children: [
                    new TextRun({
                        text: '11. Residuals',
                        ...style,
                        bold: true,
                        underline: true
                    })
                ],
                spacing: {
                    before: 240,
                    after: 120
                }
            });

            const residualsClause = new Paragraph({
                children: [
                    new TextRun({
                        text: 'Nothing in this Agreement shall be construed to limit the Receiving Party\'s right to independently develop or acquire products or services without use of the Disclosing Party\'s Confidential Information, nor shall it restrict the use of any general knowledge, skills, or experience retained in unaided memory by personnel of the Receiving Party.',
                        ...style
                    })
                ],
                spacing: {
                    before: 120,
                    after: 240,
                    line: 360
                }
            });

            doc.sections[0].children.push(heading, residualsClause);
        } catch (error) {
            console.error('Error in insertResidualsClause:', error);
            throw error;
        }
    }

    getParagraphText(paragraph) {
        return paragraph.children
            ?.map(child => child.text || '')
            .join('')
            .trim() || '';
    }
}

module.exports = DocumentService;