const path = require('path');

const { generatePermissionDocsFromGraphDocs } = require('./lib/graph-docs-parser');
const { runCommand } = require('./lib/cli');
const { createLogger } = require('./lib/logger');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_OUTPUT_DIR = path.join(ROOT_DIR, 'data');

const command = {
    name: 'refresh-docs',
    summary: 'Parse Microsoft Learn docs into permission methods, PowerShell, and code examples.',
    usage: 'refresh-docs [options]',
    options: [
        {
            name: 'output-dir',
            type: 'string',
            description: 'Directory to write the parsed doc datasets.',
            default: 'data'
        },
        {
            name: 'repo-dir',
            type: 'string',
            description: 'Reuse an existing microsoft-graph-docs checkout instead of cloning.'
        }
    ],
    run(args) {
        const log = createLogger({ scope: 'refresh-docs' });
        const outputDir = path.resolve(args['output-dir'] || DEFAULT_OUTPUT_DIR);
        const repoDir = args['repo-dir'] ? path.resolve(args['repo-dir']) : null;

        const result = generatePermissionDocsFromGraphDocs(outputDir, { repoDir });
        log.info(
            `generated api methods: permissions=${result.api.permissions} mappings=${result.api.mappings} files=${result.api.filesWithMappings}/${result.api.filesParsed}`
        );
        log.info(
            `generated powershell methods: permissions=${result.powershell.permissions} mappings=${result.powershell.mappings} files=${result.powershell.filesWithMappings}/${result.powershell.filesParsed}`
        );
        log.info(
            `generated code examples: permissions=${result.codeExamples.permissions} snippets=${result.codeExamples.snippets} files=${result.codeExamples.filesWithMappings}/${result.codeExamples.filesParsed}`
        );
        log.info(
            `generated resource docs: resources=${result.resources.resources} properties=${result.resources.propertyTables} relationships=${result.resources.relationshipTables} json=${result.resources.jsonRepresentations}`
        );
        log.info(`output dir: ${result.outputDir}`);
        return result;
    }
};

if (require.main === module) {
    runCommand(command);
}

module.exports = { command, runCli: command.run };
