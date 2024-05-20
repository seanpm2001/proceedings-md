import { spawn } from 'child_process'

export function renderLatexToMathML(latex: string) {
    return new Promise((resolve, reject) => {
        const latexmlProcess = spawn('latexml', ['--preload=amsmath', '--quiet', '-'], {
            stdio: ['pipe', 'pipe', 'pipe']
        });

        let mathml = '';
        let errors = '';

        latexmlProcess.stdout.on('data', (data) => {
            mathml += data.toString();
        });

        latexmlProcess.stderr.on('data', (data) => {
            errors += data.toString();
        });

        latexmlProcess.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(`latexml process exited with code ${code}. stderr: ${errors}`));
            } else {
                if(errors) {
                    reject(new Error(`latexml reported errors: ${errors}`))
                } else {
                    resolve(mathml);
                }
            }
        });

        latexmlProcess.on('error', (err) => {
            reject(err);
        });

        latexmlProcess.stdin.write(latex);
        latexmlProcess.stdin.end();
    });
}