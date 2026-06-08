
    
// ============================
// FECHA Y HORA AUTOMÁTICA
// ============================

const fechaSpan = document.getElementById("fecha");
const horaSpan = document.getElementById("hora");

function actualizarFechaHora() {

    const ahora = new Date();

    fechaSpan.textContent =
        ahora.toLocaleDateString("es-PE");

    horaSpan.textContent =
        ahora.toLocaleTimeString("es-PE", 
        { 
            hour: "2-digit",
            minute: "2-digit"
        }    
        );

}

actualizarFechaHora();

setInterval(actualizarFechaHora, 1000);


// ============================
// VARIABLES GLOBALES
// ============================

const contenedorJuntas =
    document.getElementById("contenedorJuntas");

const btnAgregar =
    document.getElementById("btnAgregar");

let contadorJuntas = 0;


// ============================
// AGREGAR JUNTA
// ============================

btnAgregar.addEventListener(
    "click",
    agregarJunta
);


// Primera junta automática
agregarJunta();

function agregarJunta() {

    contadorJuntas++;

    const junta = document.createElement("div");

    junta.classList.add("junta");

    junta.innerHTML = `
    
        <h2>Junta #${contadorJuntas}</h2>

        <label>Nombre de la Junta</label>

        <input
            type="text"
            class="nombreJunta"
            placeholder="Ingrese nombre de la junta"
        >

        <label>Observaciones</label>

        <textarea
            class="observacionJunta"
            placeholder="Ingrese observaciones"
        ></textarea>

        <label>Imágenes</label>

        <input
            type="file"
            class="imagenesJunta"
            multiple
            accept="image/*"
            capture="environment"
        >

        <div class="preview"></div>

        <button class="btnEliminar">
            Eliminar Junta
        </button>

    `;

    contenedorJuntas.appendChild(junta);

    configurarVistaPrevia(junta);

    configurarEliminar(junta);

}


// ============================
// ELIMINAR JUNTA
// ============================

function configurarEliminar(junta) {

    const btnEliminar =
        junta.querySelector(".btnEliminar");

    btnEliminar.addEventListener(
        "click",
        () => {

            if (
                confirm(
                    "¿Desea eliminar esta junta?"
                )
            ) {

                junta.remove();

            }

        }
    );

}


// ============================
// VISTA PREVIA DE IMÁGENES
// ============================

function configurarVistaPrevia(junta) {

    const input =
        junta.querySelector(".imagenesJunta");

    const preview =
        junta.querySelector(".preview");

    input.addEventListener(
        "change",
        () => {

            preview.innerHTML = "";

            const archivos =
                input.files;

            for (
                let i = 0;
                i < archivos.length;
                i++
            ) {

                const archivo =
                    archivos[i];

                const lector =
                    new FileReader();

                lector.onload =
                    function (e) {

                        const img =
                            document.createElement(
                                "img"
                            );

                        img.src =
                            e.target.result;

                        preview.appendChild(
                            img
                        );

                    };

                lector.readAsDataURL(
                    archivo
                );

            }

        }
    );

}

// ============================
// EXPORTAR EXCEL CON IMÁGENES
// ============================



async function exportarExcel() {

    const workbook =
        new ExcelJS.Workbook();

    const worksheet =
        workbook.addWorksheet(
            "Reporte"
        );

    worksheet.columns = [
        {
            header: "RQ",
            key: "rq",
            width: 20
        },
        {
            header: "Junta",
            key: "junta",
            width: 30
        },
        {
            header: "Observación",
            key: "obs",
            width: 40
        },
        {
            header: "Fecha",
            key: "fecha",
            width: 20
        },
        {
            header: "Hora",
            key: "hora",
            width: 20
        }
    ];

    const rq =
        document
        .getElementById("rq")
        .value;

    const juntas =
        document
        .querySelectorAll(".junta");

    let filaActual = 2;

    for (const junta of juntas) {

        const nombre =
            junta.querySelector(
                ".nombreJunta"
            ).value;

        const observacion =
            junta.querySelector(
                ".observacionJunta"
            ).value;

        worksheet.addRow({
            rq,
            junta: nombre,
            obs: observacion,
            fecha:
                document
                .getElementById("fecha")
                .textContent,
            hora:
                document
                .getElementById("hora")
                .textContent
        });

        const input =
            junta.querySelector(
                ".imagenesJunta"
            );

        const archivos =
            input.files;

        let columnaImagen = 6;

        for (
            let i = 0;
            i < archivos.length;
            i++
        ) {

            const archivo =
                archivos[i];

            const base64 =
                await convertirBase64(
                    archivo
                );

            const imageId =
                workbook.addImage({
                    base64:
                        base64,
                    extension:
                        obtenerExtension(
                            archivo.name
                        )
                });

            worksheet.addImage(
                imageId,
                {
                    tl: {
                        col:
                            columnaImagen,
                        row:
                            filaActual - 1
                    },
                    ext: {
                        width: 120,
                        height: 120
                    }
                }
            );

            worksheet.getColumn(
                columnaImagen + 1
            ).width = 20;

            columnaImagen++;

        }

        worksheet.getRow(
            filaActual
        ).height = 100;

        filaActual++;

    }

    const buffer =
        await workbook.xlsx.writeBuffer();

    saveAs(
        new Blob(
            [buffer],
            {
                type:
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        ),
        "Reporte_Juntas.xlsx"
    );

}


// ============================
// CONVERTIR IMAGEN A BASE64
// ============================

function convertirBase64(
    archivo
) {

    return new Promise(
        (resolve) => {

            const reader =
                new FileReader();

            reader.onload =
                () => {

                    resolve(
                        reader.result
                    );

                };

            reader.readAsDataURL(
                archivo
            );

        }
    );

}


// ============================
// OBTENER EXTENSIÓN
// ============================

function obtenerExtension(
    nombre
) {

    const partes =
        nombre.split(".");

    return partes[
        partes.length - 1
    ].toLowerCase();

}

// ============================
// EXPORTAR PDF
// ============================



async function exportarPDF() {

    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF();

    const rq =
        document
        .getElementById("rq")
        .value;

    const fecha =
        document
        .getElementById("fecha")
        .textContent;

    const hora =
        document
        .getElementById("hora")
        .textContent;

    let y = 20;

    pdf.setFontSize(18);

    pdf.text(
        "REPORTE DE JUNTAS",
        15,
        y
    );

    y += 15;

    pdf.setFontSize(12);

    pdf.text(
        `RQ: ${rq}`,
        15,
        y
    );

    y += 8;

    pdf.text(
        `Fecha: ${fecha}`,
        15,
        y
    );

    y += 8;

    pdf.text(
        `Hora: ${hora}`,
        15,
        y
    );

    y += 15;

    const juntas =
        document.querySelectorAll(
            ".junta"
        );

    let numero = 1;

    for (const junta of juntas) {

        if (y > 250) {

            pdf.addPage();

            y = 20;

        }

        const nombre =
            junta.querySelector(
                ".nombreJunta"
            ).value;

        const observacion =
            junta.querySelector(
                ".observacionJunta"
            ).value;

        pdf.setFontSize(14);

        pdf.text(
            `JUNTA #${numero}`,
            15,
            y
        );

        y += 10;

        pdf.setFontSize(11);

        pdf.text(
            `Nombre: ${nombre}`,
            15,
            y
        );

        y += 8;

        pdf.text(
            `Observación: ${observacion}`,
            15,
            y
        );

        y += 10;

        const input =
            junta.querySelector(
                ".imagenesJunta"
            );

        const archivos =
            input.files;

        for (
            let i = 0;
            i < archivos.length;
            i++
        ) {

            const archivo =
                archivos[i];

            const base64 =
                await convertirBase64(
                    archivo
                );

            if (y > 220) {

                pdf.addPage();

                y = 20;

            }

            pdf.addImage(
                base64,
                "JPEG",
                15,
                y,
                60,
                60
            );

            y += 70;

        }

        y += 10;

        numero++;

    }

    pdf.save(
        "Reporte_Juntas.pdf"
    );

}

// ============================
// COMPARTIR WHATSAPP
// ============================

document
.getElementById("btnWhatsapp")
.addEventListener(
    "click",
    compartirWhatsApp
);

function compartirWhatsApp() {

    const rq =
        document
        .getElementById("rq")
        .value;

    const fecha =
        document
        .getElementById("fecha")
        .textContent;

    const hora =
        document
        .getElementById("hora")
        .textContent;

    const totalJuntas =
        document
        .querySelectorAll(".junta")
        .length;

    const mensaje =
`REPORTE DE JUNTAS

RQ: ${rq}

Fecha: ${fecha}

Hora: ${hora}

Total de juntas: ${totalJuntas}`;

    const url =
        `https://wa.me/?text=${encodeURIComponent(mensaje)}`;

    window.open(
        url,
        "_blank"
    );

}


// ============================
// RENUMERAR JUNTAS
// ============================

function renumerarJuntas() {

    const juntas =
        document.querySelectorAll(
            ".junta"
        );

    juntas.forEach(
        (
            junta,
            index
        ) => {

            const titulo =
                junta.querySelector(
                    "h2"
                );

            titulo.textContent =
                `Junta #${index + 1}`;

        }
    );

}


// ============================
// REEMPLAZAR CONFIGURAR ELIMINAR
// ============================

function configurarEliminar(junta) {

    const btnEliminar =
        junta.querySelector(
            ".btnEliminar"
        );

    btnEliminar.addEventListener(
        "click",
        () => {

            const respuesta =
                confirm(
                    "¿Desea eliminar esta junta?"
                );

            if (!respuesta)
                return;

            junta.remove();

            renumerarJuntas();

        }
    );

}


// ============================
// CONTAR IMÁGENES
// ============================

function obtenerTotalImagenes() {

    let total = 0;

    document
        .querySelectorAll(
            ".imagenesJunta"
        )
        .forEach(
            (input) => {

                total +=
                    input.files.length;

            }
        );

    return total;

}


// ============================
// CONFIRMACIÓN EXCEL
// ============================

const botonExcel =
    document.getElementById(
        "btnExcel"
    );

botonExcel.addEventListener(
    "click",
    () => {

        const continuar =
            confirm(
                "¿Desea exportar a Excel?"
            );

        if (!continuar)
            return;

        exportarExcel();

    }
);


// ============================
// CONFIRMACIÓN PDF
// ============================

const botonPDF =
    document.getElementById(
        "btnPDF"
    );

botonPDF.addEventListener(
    "click",
    () => {

        const continuar =
            confirm(
                "¿Desea exportar a PDF?"
            );

        if (!continuar)
            return;

        exportarPDF();

    }
);


// ============================
// RESUMEN GENERAL
// ============================

function mostrarResumen() {

    const totalJuntas =
        document
        .querySelectorAll(
            ".junta"
        )
        .length;

    const totalImagenes =
        obtenerTotalImagenes();

    console.log(
        "Juntas:",
        totalJuntas
    );

    console.log(
        "Imágenes:",
        totalImagenes
    );

}
