const input_file = document.querySelector('#to-choose');
const output_db = document.querySelector('#output-db');
const container_choose = document.querySelector('.container-choose');
const create_sheet = document.querySelector('#create-sheet');
const output_newdb = document.querySelector('#output-newdb'); 
const download_sheet = document.querySelector('#download-sheet');
const input_name_sheet = document.querySelector('#input-name-sheet');
const container_down = document.querySelector('.container-down');
const btn_back = document.querySelector('.btn-back');
const container_down_main = document.querySelector('.container-down-main');

container_choose.addEventListener('click', input_file.click());

const db = [];

const createSpreadsheet = () => {
    output_db.innerHTML = '';

    if (db.length > 0) {
        const table = document.createElement('table');
        table.classList.add('table');
        output_db.appendChild(table);

        const headerRow = document.createElement('tr');
        table.appendChild(headerRow);
        Object.keys(db[0]).forEach(key => {
            if (key !== '__EMPTY') {
                const th = document.createElement('th');
                th.textContent = key; 
                headerRow.appendChild(th);
            }
        });

        for (let i = 0; i < db.length; i++) {
            const tr = document.createElement('tr');
            tr.classList.add('tr');
            table.appendChild(tr);

            Object.keys(db[i]).forEach(key => {
                if (key !== '__EMPTY') {
                    const td = document.createElement('td');
                    td.classList.add('td');
                    td.textContent = db[i][key]; 

                    tr.appendChild(td);
                }
            });
        }
    } else {
        alert('Arquivo em Branco ou corrompido');
    }
}

input_file.addEventListener('change', event => {
    const file = event.target.files[0]; 
    const reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result); 
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0]; 
        const sheet = workbook.Sheets[sheetName]; 
        const json = XLSX.utils.sheet_to_json(sheet, {
            raw: false,
            dateNF: 'dd/mm/yy',
            defval: '',
        }); 

        db.length = 0; 
        db.push(...json);
        createSpreadsheet();
        create_sheet.style.display = 'block';
        container_choose.style.display = 'none';
        output_db.style.display = 'block';
        output_newdb.style.display = 'block';
        btn_back.style.display = 'block';
        container_down_main.style.display = 'block';
        container_choose.textContent = 'Selecionar outro arquivo' 
        //console.log(JSON.stringify(json, null, 2));
    }

    reader.readAsArrayBuffer(file); 
});

const generateNewSpreadsheet = () => {
    const months = [
        { month: "JANEIRO", num: "01" },
        { month: "FEVEREIRO", num: "02" },
        { month: "MARÇO", num: "03" },
        { month: "ABRIL", num: "04" },
        { month: "MAIO", num: "05" },
        { month: "JUNHO", num: "06" },
        { month: "JULHO", num: "07" },
        { month: "AGOSTO", num: "08" },
        { month: "SETEMBRO", num: "09" },
        { month: "OUTUBRO", num: "10" },
        { month: "NOVEMBRO", num: "11" },
        { month: "DEZEMBRO", num: "12" }
    ];

    const getMonthName = (monthIndex) => {
        const monthNum = (monthIndex % 12) + 1;
        return months.find(monthObj => monthObj.num === monthNum.toString().padStart(2, '0')).month;
    };

    let newDataForMaturity = null;
    let day = null;

    if (db.length > 0) {
        const table = document.createElement('table');
        table.classList.add('newTable');
        output_newdb.innerHTML = '';  
        output_newdb.appendChild(table);

        const headerRow = document.createElement('tr');
        table.appendChild(headerRow);
        Object.keys(db[0]).forEach(key => {
            if (key !== '__EMPTY') {
                const th = document.createElement('th');
                th.textContent = key;
                headerRow.appendChild(th);
            }
        });

        db.forEach(row => {
            const tr = document.createElement('tr');
            tr.classList.add('tr');
            table.appendChild(tr);

            Object.keys(row).forEach(key => {
                if (key !== '__EMPTY') {
                    const td = document.createElement('td');
                    td.classList.add('td');

                    if (key === "1ª ASSEMBLEIA") {
                        newDataForMaturity = row[key].match(/\d{2}\/(\d{2})\/(\d{2})/);
                        td.textContent = row[key];
                        tr.appendChild(td);
                    } else if (key === "VENCIMENTO") {
                        day = row[key];

                        if (newDataForMaturity && newDataForMaturity.length >= 3) {
                            if (newDataForMaturity[1] == 12) {
                                td.textContent = `${day}/01/${(parseInt(newDataForMaturity[2]) + 1).toString().padStart(2, '0')}`;
                            } else {
                                td.textContent = `${day}/${(parseInt(newDataForMaturity[1]) + 1).toString().padStart(2, '0')}/${newDataForMaturity[2]}`;
                            }
                        } else {
                            td.textContent = day;
                            //console.error("Data inválida ou formato incorreto em '1ª ASSEMBLEIA'");
                        }

                        tr.appendChild(td);
                    } else if (key.includes("PARCELA")) {
                        const parcelaIndex = parseInt(key.split(" ")[0]) - 2;
                        if (newDataForMaturity && newDataForMaturity.length >= 3) {
                            const monthIndex = (parseInt(newDataForMaturity[1]) + parcelaIndex) % 12; 
                            td.textContent = getMonthName(monthIndex);
                        }
                        tr.appendChild(td);
                    } else {
                        td.textContent = row[key];
                        tr.appendChild(td);
                    }
                }
            });
        });
    } else {
        alert('Selecione um arquivo antes');
    }

    container_down.style.display = 'block'
    download_sheet.style.display = 'block'

    download_sheet.addEventListener('click', () => {
        const wb = XLSX.utils.table_to_book(document.querySelector('.newTable'), { sheet: "Sheet1" });
        if (input_name_sheet.value !== '') {
            XLSX.writeFile(wb, `${input_name_sheet.value}.xlsx`);
        }
    });
};

btn_back.addEventListener('click', () => {
    location.reload();
})

create_sheet.addEventListener('click', () => {
    generateNewSpreadsheet();
    create_sheet.style.display = 'none'
});
