// Configuración de la API de Google Sheets
const CLIENT_ID = '353952007127-8nlja4fkvr71g1mhdjash2193ud1vpt8.apps.googleusercontent.com';
const API_KEY = 'AIzaSyDzdNyVRbicnktNrk1U386-UyAXScAJny0';
const SPREADSHEET_ID = '1XXtmNhlleKC808A8EzwGUfCn_ZQqLzqYynVW-OdFZ5I';
const DISCOVERY_DOCS = ['https://sheets.googleapis.com/$discovery/rest?version=v4'];
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

// Configuración de rangos
const SHEET_NAME = 'sheet3';
let currentRange = 'B3:E'; // Rango por defecto
let lastRow = 1; // Última fila con datos

let gapiInited = false;
let gisInited = false;

// Cargar la API de Google
function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}

// Inicializar el cliente de Google
async function initializeGapiClient() {
    try {
        await gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: DISCOVERY_DOCS,
        });
        gapiInited = true;
        maybeEnableButtons();
    } catch (error) {
        console.error('Error al inicializar el cliente:', error);
    }
}

function maybeEnableButtons() {
    if (gapiInited) {
        document.getElementById('authorize-button').style.display = 'block';
    }
}

async function handleAuthClick() {
    const tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: async (resp) => {
            if (resp.error !== undefined) {
                throw resp;
            }
            document.getElementById('authorize-button').style.display = 'none';
            document.getElementById('signout-button').style.display = 'block';
            await initializeSheet();
            await listRecords();
        }
    });

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({prompt: 'consent'});
    } else {
        tokenClient.requestAccessToken({prompt: ''});
    }
}

function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
        document.getElementById('authorize-button').style.display = 'block';
        document.getElementById('signout-button').style.display = 'none';
        document.getElementById('recordsBody').innerHTML = '';
    }
}

// Inicializar la hoja de cálculo
async function initializeSheet() {
    try {
        // Verificar si la hoja existe
        const response = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID,
            ranges: [`${SHEET_NAME}!A1:E1`],
            includeGridData: true
        });

        // Si la hoja no existe, crearla
        if (!response.result.sheets || response.result.sheets.length === 0) {
            await createSheet();
        } else {
            // Obtener la última fila con datos
            const valuesResponse = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID,
                range: `${SHEET_NAME}!B3:E`
            });
            
            if (valuesResponse.result.values) {
                lastRow = valuesResponse.result.values.length + 3; // +3 porque empezamos en B3
                currentRange = `B${lastRow}:E`;
            } else {
                lastRow = 3;
                currentRange = 'B3:E';
            }
        }
    } catch (err) {
        console.error('Error al inicializar la hoja:', err);
        await createSheet();
    }
}

// Crear nueva hoja con encabezados
async function createSheet() {
    try {
        // Verificar si la hoja existe
        const sheets = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID
        });

        // Si la hoja no existe, crearla
        if (!sheets.result.sheets.some(sheet => sheet.properties.title === SHEET_NAME)) {
            await gapi.client.sheets.spreadsheets.batchUpdate({
                spreadsheetId: SPREADSHEET_ID,
                resource: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: SHEET_NAME
                            }
                        }
                    }]
                }
            });
        }

        // Agregar encabezados
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!B2:E2`,
            valueInputOption: 'RAW',
            resource: {
                values: [['Fecha y Hora', 'Máquina', 'Operador', 'Tiempo de Uso']]
            }
        });

        lastRow = 3;
        currentRange = 'B3:E';
    } catch (err) {
        console.error('Error al crear la hoja:', err);
        throw err;
    }
}

async function listRecords() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!B3:E`,
        });
        const records = response.result.values || [];
        const recordsBody = document.getElementById('recordsBody');
        recordsBody.innerHTML = '';
        
        records.forEach(record => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${record[0]}</td>
                <td>${record[1]}</td>
                <td>${record[2]}</td>
                <td>${record[3]}</td>
            `;
            recordsBody.appendChild(row);
        });
    } catch (err) {
        console.error('Error al cargar registros:', err);
        alert('Error al cargar los registros. Por favor, intente nuevamente.');
    }
}

class Timer {
    constructor(machineNumber) {
        this.machineNumber = machineNumber;
        this.hours = 0;
        this.minutes = 0;
        this.seconds = 0;
        this.milliseconds = 0;
        this.totalHours = 0;
        this.totalMinutes = 0;
        this.totalSeconds = 0;
        this.totalMilliseconds = 0;
        this.cycles = 0;
        this.isRunning = false;
        this.timer = null;
        
        this.initializeElements();
        this.attachEventListeners();
    }

    initializeElements() {
        this.hoursElement = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .hours`);
        this.minutesElement = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .minutes`);
        this.secondsElement = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .seconds`);
        this.millisecondsElement = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .milliseconds`);
        this.startBtn = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .startBtn`);
        this.stopBtn = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .stopBtn`);
        this.resetBtn = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .resetBtn`);
        this.saveBtn = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .saveBtn`);
        this.operatorInput = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .operatorName`);
        this.neonIndicator = document.querySelector(`.machine-panel:nth-child(${this.machineNumber}) .neon-indicator`);
        
        this.totalHoursElement = document.querySelector(`.total-time-card:nth-child(${this.machineNumber}) .total-hours`);
        this.totalMinutesElement = document.querySelector(`.total-time-card:nth-child(${this.machineNumber}) .total-minutes`);
        this.totalSecondsElement = document.querySelector(`.total-time-card:nth-child(${this.machineNumber}) .total-seconds`);
        this.totalMillisecondsElement = document.querySelector(`.total-time-card:nth-child(${this.machineNumber}) .total-milliseconds`);
        this.cycleCountElement = document.querySelector(`.total-time-card:nth-child(${this.machineNumber}) .cycle-count`);
    }

    attachEventListeners() {
        this.startBtn.addEventListener('click', () => this.start());
        this.stopBtn.addEventListener('click', () => this.stop());
        this.resetBtn.addEventListener('click', () => this.reset());
        this.saveBtn.addEventListener('click', () => this.saveRecord());
    }

    updateDisplay() {
        this.hoursElement.textContent = this.hours.toString().padStart(2, '0');
        this.minutesElement.textContent = this.minutes.toString().padStart(2, '0');
        this.secondsElement.textContent = this.seconds.toString().padStart(2, '0');
        this.millisecondsElement.textContent = this.milliseconds.toString().padStart(3, '0');
    }

    updateTotalDisplay() {
        this.totalHoursElement.textContent = this.totalHours.toString().padStart(2, '0');
        this.totalMinutesElement.textContent = this.totalMinutes.toString().padStart(2, '0');
        this.totalSecondsElement.textContent = this.totalSeconds.toString().padStart(2, '0');
        this.totalMillisecondsElement.textContent = this.totalMilliseconds.toString().padStart(3, '0');
    }

    addToTotalTime() {
        this.totalMilliseconds += this.milliseconds;
        this.totalSeconds += this.seconds;
        this.totalMinutes += this.minutes;
        this.totalHours += this.hours;

        // Ajustar el tiempo total
        while (this.totalMilliseconds >= 1000) {
            this.totalMilliseconds -= 1000;
            this.totalSeconds++;
        }
        while (this.totalSeconds >= 60) {
            this.totalSeconds -= 60;
            this.totalMinutes++;
        }
        while (this.totalMinutes >= 60) {
            this.totalMinutes -= 60;
            this.totalHours++;
        }

        this.updateTotalDisplay();
    }

    start() {
        if (!this.isRunning) {
            this.isRunning = true;
            this.neonIndicator.classList.add('active');
            this.timer = setInterval(() => {
                this.milliseconds += 10;
                if (this.milliseconds === 1000) {
                    this.milliseconds = 0;
                    this.seconds++;
                    if (this.seconds === 60) {
                        this.seconds = 0;
                        this.minutes++;
                        if (this.minutes === 60) {
                            this.minutes = 0;
                            this.hours++;
                        }
                    }
                }
                this.updateDisplay();
            }, 10);
        }
    }

    stop() {
        if (this.isRunning) {
            this.isRunning = false;
            this.neonIndicator.classList.remove('active');
            clearInterval(this.timer);
            this.cycles++;
            this.updateCycleCount();
        }
    }

    reset() {
        this.stop();
        this.hours = 0;
        this.minutes = 0;
        this.seconds = 0;
        this.milliseconds = 0;
        this.updateDisplay();
    }

    updateCycleCount() {
        this.cycleCountElement.textContent = this.cycles;
    }

    async saveRecord() {
        if (!this.operatorInput.value) {
            alert('Por favor, ingrese el nombre del operador');
            return;
        }

        // Verificar si el usuario está autenticado
        if (!gapi.client.getToken()) {
            alert('Por favor, inicie sesión con Google para guardar registros');
            return;
        }

        const now = new Date();
        const dateTime = now.toLocaleString();
        const timeUsed = `${this.hours.toString().padStart(2, '0')}:${this.minutes.toString().padStart(2, '0')}:${this.seconds.toString().padStart(2, '0')}.${this.milliseconds.toString().padStart(3, '0')}`;

        try {
            // Verificar si la hoja existe antes de guardar
            await initializeSheet();

            // Guardar el registro
            const response = await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: SPREADSHEET_ID,
                range: `${SHEET_NAME}!${currentRange}`,
                valueInputOption: 'RAW',
                resource: {
                    values: [[dateTime, `Máquina ${this.machineNumber}`, this.operatorInput.value, timeUsed]]
                }
            });

            if (response.status === 200) {
                // Actualizar la tabla local
                const newRow = document.createElement('tr');
                newRow.innerHTML = `
                    <td>${dateTime}</td>
                    <td>Máquina ${this.machineNumber}</td>
                    <td>${this.operatorInput.value}</td>
                    <td>${timeUsed}</td>
                `;
                document.getElementById('recordsBody').appendChild(newRow);

                // Actualizar el tiempo total
                this.addToTotalTime();
                
                // Resetear el temporizador
                this.reset();
                
                // Actualizar el rango para el próximo registro
                lastRow++;
                currentRange = `B${lastRow}:E`;
                
                alert('Registro guardado exitosamente');
            } else {
                throw new Error('Error al guardar el registro');
            }
        } catch (err) {
            console.error('Error al guardar registro:', err);
            alert('Error al guardar el registro. Por favor, intente nuevamente.');
        }
    }
}

// Inicializar la API de Google cuando se carga la página
window.onload = gapiLoaded;

// Asignar los manejadores de eventos
document.getElementById('authorize-button').onclick = handleAuthClick;
document.getElementById('signout-button').onclick = handleSignoutClick;

// Crear instancias de Timer para cada máquina
const timer1 = new Timer(1);
const timer2 = new Timer(2);
const timer3 = new Timer(3);
const timer4 = new Timer(4);

function updateClock() {
    const now = new Date();
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const seconds = now.getSeconds();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    const formattedHours = hours % 12 || 12;
    const formattedMinutes = minutes.toString().padStart(2, '0');
    const formattedSeconds = seconds.toString().padStart(2, '0');
    
    document.querySelector('.hours').textContent = formattedHours;
    document.querySelector('.minutes').textContent = formattedMinutes;
    document.querySelector('.seconds').textContent = formattedSeconds;
    document.querySelector('.ampm').textContent = ampm;
}

// Inicializar el reloj
updateClock();
setInterval(updateClock, 1000); 