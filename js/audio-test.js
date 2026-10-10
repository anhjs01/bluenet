let recorder = null;
let chunks = [];
let stream = null;
let blob = null;

let ctx = null;
let source = null;
let analyser = null;
let raf = 0;

let timer = 0;
let startedAt = 0;

let selectedInputId = "";
let selectedOutputId = "";

let deviceChangeHandler = null;


function stopVisual() {
    cancelAnimationFrame(raf);

    raf = 0;

    source?.disconnect();
    analyser?.disconnect();

    ctx?.close().catch(() => {});

    source = null;
    analyser = null;
    ctx = null;
}


function formatTime(ms) {
    const seconds = Math.floor(ms / 1000);

    const minutes = Math.floor(seconds / 60);

    const remainingSeconds = seconds % 60;

    return (
        String(minutes).padStart(2, "0") +
        ":" +
        String(remainingSeconds).padStart(2, "0")
    );
}


function stopTimer(box) {
    clearInterval(timer);

    timer = 0;

    const elapsed = box.querySelector("#audioTimer");

    if (elapsed && startedAt) {
        elapsed.textContent =
            formatTime(Date.now() - startedAt);
    }
}


function startTimer(box) {
    startedAt = Date.now();

    const elapsed =
        box.querySelector("#audioTimer");

    if (elapsed) {
        elapsed.textContent = "00:00";
    }

    clearInterval(timer);

    timer = setInterval(() => {
        if (elapsed) {
            elapsed.textContent =
                formatTime(Date.now() - startedAt);
        }
    }, 250);
}


function draw(box) {
    const wave = box.querySelector(".audio-wave");
    const dot = box.querySelector(".audio-dot");
    if (!analyser || !wave || !dot) return;

    const frequencyData = new Uint8Array(analyser.frequencyBinCount);
    const timeData = new Uint8Array(analyser.fftSize);
    const bars = [...wave.querySelectorAll("i")];

    const loop = () => {
        if (!analyser || !wave.isConnected) return;
        analyser.getByteFrequencyData(frequencyData);
        analyser.getByteTimeDomainData(timeData);

        let sum = 0;
        for (const value of timeData) {
            const sample = (value - 128) / 128;
            sum += sample * sample;
        }
        const rms = Math.sqrt(sum / timeData.length);
        const peak = frequencyData.reduce((max, value) => Math.max(max, value), 0);
        const hasSignal = rms > 0.0045 && peak > 5;
        dot.classList.toggle("live", hasSignal);

        // Frecuencias distribuidas linealmente; ganancia visual elevada para voz suave.
        const nyquist = (ctx?.sampleRate || 48000) / 2;
        const maxBin = Math.min(frequencyData.length - 1, Math.floor(24000 / nyquist * frequencyData.length));
        bars.forEach((bar, index) => {
            const from = Math.floor(index * maxBin / bars.length);
            const to = Math.max(from + 1, Math.floor((index + 1) * maxBin / bars.length));
            let energy = 0;
            for (let bin = from; bin < to; bin++) energy = Math.max(energy, frequencyData[bin] || 0);
            const normalized = hasSignal ? Math.max(0, (energy - 4) / 251) : 0;
            const boosted = Math.pow(normalized, 0.62);
            bar.style.height = (3 + boosted * 68) + "px";
            bar.style.opacity = hasSignal ? String(0.48 + boosted * 0.52) : "0.28";
        });
        raf = requestAnimationFrame(loop);
    };
    loop();
}

async function getDevices() {
    if (
        !navigator.mediaDevices?.enumerateDevices
    ) {
        return [];
    }

    try {
        return await navigator.mediaDevices.enumerateDevices();
    } catch {
        return [];
    }
}


function getDeviceLabel(
    device,
    index,
    type
) {
    if (device.label) {
        return device.label;
    }

    return type === "input"
        ? `Micrófono ${index + 1}`
        : `Salida de audio ${index + 1}`;
}


function updateSelectedLabels(box) {
    const input =
        box.querySelector("#audioInput");

    const output =
        box.querySelector("#audioOutput");

    const inputName =
        box.querySelector("#selectedInputName");

    const outputName =
        box.querySelector("#selectedOutputName");

    const activeInput =
        input?.options[input.selectedIndex];

    const activeOutput =
        output?.options[output.selectedIndex];

    if (inputName) {
        inputName.textContent =
            activeInput?.textContent ||
            "No seleccionado";
    }

    if (outputName) {
        outputName.textContent =
            activeOutput?.textContent ||
            "Salida predeterminada";
    }
}


function setTestStatus(
    box,
    text
) {
    const status =
        box.querySelector("#audioState");

    if (status) {
        status.textContent = text;
    }
}


async function refreshDeviceSelectors(box) {
    const inputSelect =
        box.querySelector("#audioInput");

    const outputSelect =
        box.querySelector("#audioOutput");

    if (
        !inputSelect ||
        !outputSelect
    ) {
        return;
    }

    const devices =
        await getDevices();

    const inputs =
        devices.filter(
            device =>
                device.kind === "audioinput"
        );

    const outputs =
        devices.filter(
            device =>
                device.kind === "audiooutput"
        );

    inputSelect.innerHTML = "";

    outputSelect.innerHTML = "";


    if (!inputs.length) {
        inputSelect.innerHTML =
            `
            <option value="">
                No se detectaron micrófonos
            </option>
            `;
    } else {
        inputs.forEach(
            (device, index) => {

                const option =
                    document.createElement(
                        "option"
                    );

                option.value =
                    device.deviceId;

                option.textContent =
                    getDeviceLabel(
                        device,
                        index,
                        "input"
                    );

                inputSelect.appendChild(
                    option
                );
            }
        );
    }


    if (!outputs.length) {
        outputSelect.innerHTML =
            `
            <option value="">
                Salida predeterminada del navegador
            </option>
            `;
    } else {
        outputs.forEach(
            (device, index) => {

                const option =
                    document.createElement(
                        "option"
                    );

                option.value =
                    device.deviceId;

                option.textContent =
                    getDeviceLabel(
                        device,
                        index,
                        "output"
                    );

                outputSelect.appendChild(
                    option
                );
            }
        );
    }


    if (
        selectedInputId &&
        inputs.some(
            device =>
                device.deviceId ===
                selectedInputId
        )
    ) {
        inputSelect.value =
            selectedInputId;

    } else if (inputs[0]) {

        selectedInputId =
            inputs[0].deviceId;

        inputSelect.value =
            selectedInputId;
    }


    if (
        selectedOutputId &&
        outputs.some(
            device =>
                device.deviceId ===
                selectedOutputId
        )
    ) {
        outputSelect.value =
            selectedOutputId;

    } else if (outputs[0]) {

        selectedOutputId =
            outputs[0].deviceId;

        outputSelect.value =
            selectedOutputId;
    }


    updateSelectedLabels(box);
}


async function selectInput(box) {
    const inputSelect =
        box.querySelector("#audioInput");

    const deviceId =
        inputSelect?.value;

    if (!deviceId) {
        return;
    }

    selectedInputId =
        deviceId;

    try {

        stream?.getTracks()
            .forEach(
                track =>
                    track.stop()
            );

        stopVisual();

        stream =
            await navigator.mediaDevices
                .getUserMedia({
                    audio: {
                        deviceId: {
                            exact: deviceId
                        }
                    }
                });


        const AudioContext =
            window.AudioContext ||
            window.webkitAudioContext;


        if (AudioContext) {

            ctx =
                new AudioContext();

            source =
                ctx.createMediaStreamSource(
                    stream
                );

            analyser =
                ctx.createAnalyser();

            analyser.fftSize = 2048;
            analyser.smoothingTimeConstant = 0.72;

            source.connect(
                analyser
            );

            draw(box);
        }


        updateSelectedLabels(box);

        setTestStatus(
            box,
            "Listo para grabar"
        );

    } catch {

        setTestStatus(
            box,
            "No se pudo utilizar este micrófono"
        );

        alert(
            "No se pudo utilizar la diadema seleccionada. " +
            "Comprueba que esté conectada y que Chrome tenga permiso para utilizarla."
        );
    }
}


async function selectOutput(box) {
    const outputSelect =
        box.querySelector("#audioOutput");

    const player =
        box.querySelector("#player");

    const deviceId =
        outputSelect?.value;

    if (
        !deviceId ||
        !player
    ) {
        return;
    }

    selectedOutputId =
        deviceId;

    updateSelectedLabels(box);


    if (
        typeof player.setSinkId ===
        "function"
    ) {

        try {

            await player.setSinkId(
                deviceId
            );

            const outputName =
                box.querySelector(
                    "#selectedOutputName"
                )?.textContent ||
                "salida seleccionada";

            const state =
                box.querySelector(
                    "#outputState"
                );

            if (state) {
                state.textContent =
                    `Salida activa: ${outputName}`;
            }

        } catch {

            const state =
                box.querySelector(
                    "#outputState"
                );

            if (state) {
                state.textContent =
                    "No se pudo cambiar la salida. " +
                    "Se utilizará la salida predeterminada.";
            }
        }
    }
}


function renderControls(box) {

    box.innerHTML = `

        <section class="audio-music-panel" aria-label="Prueba de parlantes">
            <div class="audio-music-icon" aria-hidden="true">♫</div>
            <div class="audio-music-copy">
                <span class="audio-music-eyebrow">PRUEBA DE SALIDA</span>
                <strong>Comprueba tus parlantes</strong>
                <p>Reproduce música local y escucha si el sonido sale con claridad.</p>
            </div>
            <div class="audio-music-controls">
                <button id="speakerMusicBtn" type="button" class="audio-music-button">
                    <span aria-hidden="true">▶</span> Reproducir música
                </button>
                <button id="speakerNextBtn" type="button" class="audio-music-next" aria-label="Pasar a la siguiente canción" title="Siguiente canción">
                    <span aria-hidden="true">⏭</span> Siguiente
                </button>
            </div>
            <div class="audio-volume-control">
                <label for="speakerMusicVolume"><span aria-hidden="true">🔊</span> Volumen <output id="speakerMusicVolumeValue" for="speakerMusicVolume">70%</output></label>
                <input id="speakerMusicVolume" type="range" min="0" max="1" step="0.05" value="0.7" aria-label="Volumen de la música">
            </div>
        </section>
        <audio id="speakerMusicPlayer" preload="none"></audio>

        <div class="audio-device-panel">

            <div class="audio-device-field">

                <label for="audioInput">
                    🎤 Micrófono
                </label>

                <select id="audioInput">
                    <option>
                        Detectando dispositivos…
                    </option>
                </select>

            </div>


            <div class="audio-device-field">

                <label for="audioOutput">
                    🔊 Salida
                </label>

                <select id="audioOutput">
                    <option>
                        Detectando dispositivos…
                    </option>
                </select>

            </div>

        </div>


        <div class="audio-live">

            <span class="audio-dot"></span>

            <div class="audio-live-copy">

                <strong id="audioState">
                    Listo para probar
                </strong>

                <span id="audioTimer">
                    00:00
                </span>

            </div>


            <div class="audio-spectrum">
                <div class="audio-frequency-labels"><span>0 Hz</span><span>6 kHz</span><span>12 kHz</span><span>18 kHz</span><span>24 kHz</span></div>
                <div class="audio-wave">
                    ${Array.from({ length: 64 }, () => "<i></i>").join("")}
                </div>
            </div>

        </div>


        <div class="audio-main-action">

            <button
                id="rec"
                type="button"
                class="primary audio-record-button"
            >
                🎙️ Iniciar grabación
            </button>


            <button
                id="del"
                type="button"
                class="audio-delete"
                aria-label="Eliminar grabación"
                title="Eliminar grabación"
            >
                ×
            </button>

        </div>

        <audio
            id="player"
            controls
            class="full audio-player"
        ></audio>


        <div
            id="outputState"
            class="audio-device-status"
        ></div>

    `;
}


export async function audioTest(
    box,
    cb = () => {}
) {

    try {

        if (
            !navigator.mediaDevices?.getUserMedia
        ) {
            throw new Error(
                "El navegador no permite acceder al audio."
            );
        }


        stream =
            await navigator.mediaDevices
                .getUserMedia({
                    audio: true
                });


        renderControls(box);


        const AudioContext =
            window.AudioContext ||
            window.webkitAudioContext;


        if (AudioContext) {

            ctx =
                new AudioContext();

            source =
                ctx.createMediaStreamSource(
                    stream
                );

            analyser =
                ctx.createAnalyser();

            analyser.fftSize = 2048;
            analyser.smoothingTimeConstant = 0.72;

            source.connect(
                analyser
            );

            draw(box);
        }


        await refreshDeviceSelectors(
            box
        );


        const inputSelect =
            box.querySelector(
                "#audioInput"
            );

        const outputSelect =
            box.querySelector(
                "#audioOutput"
            );

        const rec =
            box.querySelector("#rec");

        const del =
            box.querySelector("#del");

        const player =
            box.querySelector("#player");

        const speakerMusicBtn = box.querySelector("#speakerMusicBtn");
        const speakerNextBtn = box.querySelector("#speakerNextBtn");
        const speakerMusicPlayer = box.querySelector("#speakerMusicPlayer");
        const speakerMusicVolume = box.querySelector("#speakerMusicVolume");
        const speakerMusicVolumeValue = box.querySelector("#speakerMusicVolumeValue");
        speakerMusicPlayer.volume = Number(speakerMusicVolume.value);
        speakerMusicVolume.oninput = () => {
            speakerMusicPlayer.volume = Number(speakerMusicVolume.value);
            speakerMusicVolumeValue.textContent = Math.round(speakerMusicPlayer.volume * 100) + "%";
        };
        const speakerTracks = [
            "./audio/Mac%20Miller%20-%20Congratulations%20feat.%20Bilal%20(No%20Intro)%20-%20J.mp3",
            "./audio/Mac%20Miller%20_%20Congratulations%20piano%20cover%20-%20proly_dreamin.mp3"
        ];
        let currentSpeakerTrack = -1;

        const playRandomSpeakerTrack = async () => {
            if (speakerTracks.length > 1) {
                let nextTrack = currentSpeakerTrack;
                while (nextTrack === currentSpeakerTrack) {
                    nextTrack = Math.floor(Math.random() * speakerTracks.length);
                }
                currentSpeakerTrack = nextTrack;
            } else {
                currentSpeakerTrack = 0;
            }
            speakerMusicPlayer.src = speakerTracks[currentSpeakerTrack];
            try {
                await speakerMusicPlayer.play();
                speakerMusicBtn.innerHTML = '<span aria-hidden="true">Ⅱ</span> Pausar música';
            } catch {
                speakerMusicBtn.innerHTML = '<span aria-hidden="true">▶</span> Reproducir música';
                setTestStatus(box, "No se pudo reproducir la música. Vuelve a intentarlo.");
            }
        };

        speakerMusicBtn.onclick = async () => {
            if (!speakerMusicPlayer.paused) {
                speakerMusicPlayer.pause();
                speakerMusicBtn.innerHTML = '<span aria-hidden="true">▶</span> Reproducir música';
                return;
            }
            if (!speakerMusicPlayer.src) {
                await playRandomSpeakerTrack();
            } else {
                try {
                    await speakerMusicPlayer.play();
                    speakerMusicBtn.innerHTML = '<span aria-hidden="true">Ⅱ</span> Pausar música';
                } catch {
                    setTestStatus(box, "No se pudo reproducir la música. Vuelve a intentarlo.");
                }
            }
        };

        speakerNextBtn.onclick = async () => {
            const wasPlaying = !speakerMusicPlayer.paused;
            speakerMusicPlayer.pause();
            await playRandomSpeakerTrack();
            if (!wasPlaying && !speakerMusicPlayer.paused) {
                speakerMusicPlayer.pause();
                speakerMusicBtn.innerHTML = '<span aria-hidden="true">▶</span> Reproducir música';
            }
        };

        speakerMusicPlayer.onended = () => {
            playRandomSpeakerTrack();
        };

        const state =
            box.querySelector(
                "#audioState"
            );

        const dot =
            box.querySelector(
                ".audio-dot"
            );


        inputSelect.onchange =
            async () => {

                updateSelectedLabels(
                    box
                );

                await selectInput(
                    box
                );
            };


        outputSelect.onchange =
            async () => {

                updateSelectedLabels(
                    box
                );

                await selectOutput(
                    box
                );
            };


        deviceChangeHandler =
            async () => {

                await refreshDeviceSelectors(
                    box
                );

                if (
                    selectedInputId &&
                    !box.querySelector(
                        "#audioInput"
                    )?.value
                ) {

                    setTestStatus(
                        box,
                        "La diadema seleccionada ya no está conectada"
                    );
                }
            };


        navigator.mediaDevices
            .addEventListener?.(
                "devicechange",
                deviceChangeHandler
            );


        rec.onclick =
            async () => {

                if (
                    recorder?.state ===
                    "recording"
                ) {

                    recorder.stop();

                    return;
                }


                if (
                    !stream ||
                    !stream.active
                ) {

                    await selectInput(
                        box
                    );
                }


                if (
                    !stream?.active
                ) {
                    return;
                }


                chunks = [];

                blob = null;


                recorder =
                    new MediaRecorder(
                        stream
                    );


                recorder.ondataavailable =
                    event => {

                        if (
                            event.data.size
                        ) {
                            chunks.push(
                                event.data
                            );
                        }
                    };


                recorder.onstop =
                    () => {

                        stopTimer(
                            box
                        );


                        blob =
                            new Blob(
                                chunks,
                                {
                                    type:
                                        recorder.mimeType ||
                                        "audio/webm"
                                }
                            );


                        player.src =
                            URL.createObjectURL(
                                blob
                            );


                        rec.textContent =
                            "🎙️ Grabar de nuevo";


                        state.textContent =
                            "Grabación lista para reproducir";


                        dot.classList.remove(
                            "live"
                        );


                        cb("Sí");
                    };


                recorder.start();


                startTimer(
                    box
                );


                rec.textContent =
                    "⏹ Parar grabación";


                state.textContent =
                    "Grabando… habla para comprobar el micrófono";


                dot.classList.add(
                    "live"
                );


                setTimeout(
                    () => {

                        if (
                            recorder?.state ===
                            "recording"
                        ) {
                            recorder.stop();
                        }

                    },
                    45000
                );
            };


        del.onclick =
            () => {

                if (
                    recorder?.state ===
                    "recording"
                ) {
                    recorder.stop();
                }


                blob = null;


                player.removeAttribute(
                    "src"
                );

                player.load();


                stopTimer(
                    box
                );


                rec.textContent =
                    "🎙️ Iniciar grabación";


                state.textContent =
                    "Listo para probar";


                dot.classList.remove(
                    "live"
                );


                cb("No");
            };


        stream.getTracks()
            .forEach(
                track => {

                    track.addEventListener(
                        "ended",
                        () => {

                            state.textContent =
                                "⚠️ La diadema fue desconectada";

                            dot.classList.remove(
                                "live"
                            );
                        }
                    );
                }
            );

    } catch {

        box.innerHTML =
            `
            <span class="muted">
                No se pudo iniciar el test de audio.
                Puedes omitirlo.
            </span>
            `;

        cb("Omitido");
    }
}


export function cleanupAudio() {

    const speakerMusicPlayer = document.querySelector("#speakerMusicPlayer");
    if (speakerMusicPlayer) {
        speakerMusicPlayer.pause();
        speakerMusicPlayer.removeAttribute("src");
        speakerMusicPlayer.load();
    }

    clearInterval(timer);

    timer = 0;


    if (
        recorder?.state ===
        "recording"
    ) {
        recorder.stop();
    }


    stopVisual();


    stream?.getTracks()
        .forEach(
            track =>
                track.stop()
        );


    if (
        navigator.mediaDevices &&
        deviceChangeHandler
    ) {

        navigator.mediaDevices
            .removeEventListener(
                "devicechange",
                deviceChangeHandler
            );
    }


    deviceChangeHandler =
        null;

    stream = null;

    recorder = null;

    blob = null;

    chunks = [];

    selectedInputId = "";

    selectedOutputId = "";
}