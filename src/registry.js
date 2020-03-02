const resumes = {};

function load() {
    const requireComponent = require.context('../resume', true, /.yml$/);

    requireComponent.keys().forEach(fileName => {
        const key = fileName.split('.')[1].substr(1);

        resumes[key] = requireComponent(fileName).PERSON;
    });
}

export default resumes;
export { load };