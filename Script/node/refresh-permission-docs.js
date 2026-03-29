const path = require('path');

const { generatePermissionDocsFromGraphDocs } = require('./lib/graph-docs-parser');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_OUTPUT_DIR = path.join(ROOT_DIR, 'data');

function parseArgs(argv) {
    const args = {};
    for (let index = 0; index < argv.length; index += 1) {
        const part = argv[index];
        if (!part.startsWith('--')) {
            continue;
        }

        const key = part.slice(2);
        const next = argv[index + 1];
        if (!next || next.startsWith('--')) {
            args[key] = true;
        } else {
            args[key] = next;
            index += 1;
        }
    }
    return args;
}

function runCli() {
    const args = parseArgs(process.argv.slice(2));
    const outputDir = path.resolve(args['output-dir'] || DEFAULT_OUTPUT_DIR);
    const repoDir = args['repo-dir'] ? path.resolve(args['repo-dir']) : null;

    const result = generatePermissionDocsFromGraphDocs(outputDir, { repoDir });
    console.log(`generated api methods: permissions=${result.api.permissions} mappings=${result.api.mappings} files=${result.api.filesWithMappings}/${result.api.filesParsed}`);
    console.log(`generated powershell methods: permissions=${result.powershell.permissions} mappings=${result.powershell.mappings} files=${result.powershell.filesWithMappings}/${result.powershell.filesParsed}`);
    console.log(`generated code examples: permissions=${result.codeExamples.permissions} snippets=${result.codeExamples.snippets} files=${result.codeExamples.filesWithMappings}/${result.codeExamples.filesParsed}`);
    console.log(`generated resource docs: resources=${result.resources.resources} properties=${result.resources.propertyTables} relationships=${result.resources.relationshipTables} json=${result.resources.jsonRepresentations}`);
    console.log(`output dir: ${result.outputDir}`);
}

if (require.main === module) {
    runCli();
}

module.exports = { runCli };
