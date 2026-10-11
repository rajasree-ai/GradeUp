const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const dropText = document.getElementById('dropText');

if (dropZone && fileInput) {
    dropZone.addEventListener('click', () => fileInput.click());

    ['dragenter', 'dragover'].forEach(evt =>
        dropZone.addEventListener(evt, e => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        })
    );

    ['dragleave', 'drop'].forEach(evt =>
        dropZone.addEventListener(evt, e => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        })
    );

    dropZone.addEventListener('drop', e => {
        const files = e.dataTransfer.files;
        if (files.length) {
            fileInput.files = files;
            if (dropText) dropText.textContent = `📄 ${files[0].name}`;
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length && dropText) {
            dropText.textContent = `📄 ${fileInput.files[0].name}`;
        }
    });
}

setTimeout(() => {
    document.querySelectorAll('.flash').forEach(el => {
        el.style.transition = 'opacity .4s, transform .4s';
        el.style.opacity = '0';
        el.style.transform = 'translateY(-8px)';
        setTimeout(() => el.remove(), 400);
    });
}, 4000);