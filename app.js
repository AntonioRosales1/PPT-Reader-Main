document.getElementById('pptFile').addEventListener('change', handleFileSelect);
document.getElementById('playButton').addEventListener('click', playReading);
document.getElementById('pauseButton').addEventListener('click', pauseReading);
document.getElementById('stopButton').addEventListener('click', stopReading);
document.getElementById('rewindButton').addEventListener('click', rewindReading);
document.getElementById('fastForwardButton').addEventListener('click', fastForwardReading);
document.getElementById('speedRange').addEventListener('input', updateSpeed);
document.getElementById('voiceSelect').addEventListener('change', updateVoice);

let extractedText = '';
let speech = null;
let speechIndex = 0;
let isPlaying = false;
let voices = [];
let currentVoice = null;
let speakingRate = 1; // Default speech rate

// Load available voices from the Web Speech API
function loadVoices() {
    voices = window.speechSynthesis.getVoices();
    const voiceSelect = document.getElementById('voiceSelect');
    voiceSelect.innerHTML = ''; // Clear existing options

    // Add voices to the dropdown menu
    voices.forEach((voice, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${voice.name} (${voice.lang})`;
        voiceSelect.appendChild(option);
    });

    // Select the first voice as default
    if (voices.length > 0) {
        currentVoice = voices[0];
    }
}

speechSynthesis.onvoiceschanged = loadVoices;

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;

            JSZip.loadAsync(arrayBuffer).then(function(zip) {
                const slidePromises = [];
                zip.forEach(function (relativePath, zipEntry) {
                    if (relativePath.match(/ppt\/slides\/slide\d+\.xml/)) {
                        slidePromises.push(zipEntry.async("text"));
                    }
                });

                Promise.all(slidePromises).then(function(slideContents) {
                    extractedText = '';
                    slideContents.forEach(function(content) {
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(content, "application/xml");
                        const textElements = xmlDoc.getElementsByTagName("a:t");
                        for (let i = 0; i < textElements.length; i++) {
                            extractedText += textElements[i].textContent + ' ';
                        }
                    });
                    renderText();
                    enableButtons();
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Please upload a valid .pptx file');
    }
}

function renderText() {
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = '';

    const words = extractedText.trim().split(/\s+/);
    words.forEach((word, index) => {
        const span = document.createElement('span');
        span.textContent = word + ' ';
        span.dataset.index = index;
        span.addEventListener('dblclick', () => startReadingFromIndex(index));
        outputDiv.appendChild(span);
    });
}

function startReadingFromIndex(index) {
    speechIndex = index;
    playReading();
}

function playReading() {
    if (!isPlaying) {
        isPlaying = true;
        speech = new SpeechSynthesisUtterance();
        speech.text = extractedText.split(/\s+/).slice(speechIndex).join(' ');
        speech.rate = speakingRate;
        speech.voice = currentVoice; // Set the selected voice
        speechSynthesis.speak(speech);

        speech.onboundary = function(event) {
            highlightWord(event.charIndex);
        };

        speech.onend = function() {
            resetControls();
        };
    }
}

function highlightWord(startIndex) {
    const words = extractedText.trim().split(/\s+/);
    const outputDiv = document.getElementById('output');
    const spans = outputDiv.getElementsByTagName('span');

    let wordPosition = 0;
    for (let i = 0; i < words.length; i++) {
        if (wordPosition >= startIndex) {
            spans[i].classList.add('highlight');
            speechIndex = i;
            break;
        }
        spans[i].classList.remove('highlight');
        wordPosition += words[i].length + 1; // +1 for space
    }
}

function pauseReading() {
    if (isPlaying) {
        speechSynthesis.pause();
        isPlaying = false;
    }
}

function stopReading() {
    speechSynthesis.cancel();
    resetControls();
}

function rewindReading() {
    speechIndex = Math.max(0, speechIndex - 10); // Go back 10 words
    playReading();
}

function fastForwardReading() {
    const totalWords = extractedText.split(/\s+/).length;
    speechIndex = Math.min(totalWords, speechIndex + 10); // Move forward 10 words
    playReading();
}

function updateSpeed() {
    speakingRate = document.getElementById('speedRange').value;
    if (isPlaying) {
        stopReading();
        playReading(); // Restart with updated speed
    }
}

function updateVoice() {
    const selectedVoiceIndex = document.getElementById('voiceSelect').value;
    currentVoice = voices[selectedVoiceIndex];
    if (isPlaying) {
        stopReading();
        playReading(); // Restart with updated voice
    }
}

function resetControls() {
    isPlaying = false;
    speechIndex = 0;
    document.querySelectorAll('.highlight').forEach(span => {
        span.classList.remove('highlight');
    });
}

function enableButtons() {
    document.getElementById('playButton').disabled = false;
    document.getElementById('pauseButton').disabled = false;
    document.getElementById('stopButton').disabled = false;
    document.getElementById('rewindButton').disabled = false;
    document.getElementById('fastForwardButton').disabled = false;
}
