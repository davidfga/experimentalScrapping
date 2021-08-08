const puppeteer = require('puppeteer');
const Excel = require('exceljs');
const internetAvailable = require("internet-available");
const NUMBEROFPAGES = 254886; // Each page has 4 row's
const CONSULTNAME = '%';
const ROWLENGTHTOSAVE = 400; // It must be multipler of 4 
const CURRENTPAGE = 6200; //aspxGVPagerOnClick('grdResultados','PN6115')
const CURRENTWORKBOOK = 62;

function pressAnyKey(msg = 'Press any key to continue') {
    return new Promise((resolve) => {
        console.log(msg || 'Press any key to continue');
        process.stdin.setRawMode(true);
        process.stdin.resume();
        process.stdin.on('data', () => {
            process.stdin.destroy();
            resolve();
        });
    });
}

//Start
(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
    })
    const page = await browser.newPage()
    
    try {
        //Search
        console.log('connecting to website...')
        await page.goto('https://etb.com/paginasblancas/')
        await page.type('#txtNombreResl_I',CONSULTNAME)
        await page.click('#rbLista_RB0_I_D')
        await page.click('#cbLugarResl_I')
        await page.keyboard.press('ArrowUp')
        await page.keyboard.press('ArrowUp')
        await page.keyboard.press('ArrowUp')
        await page.keyboard.press('Enter')
        
        console.log('searching...')
        await page.click('#btnEncontrar_CD')

        //Pause While I search the currentpage
        await pressAnyKey()
        
        //Initial parameters
        let data = []
        let workbookNumber = CURRENTWORKBOOK
        
        for(let i = CURRENTPAGE; i <= (NUMBEROFPAGES -1 ) ; i++ ){
            console.group('Page ' + (i+1) + ' of ' + NUMBEROFPAGES)
            
            //Get data
            console.log('- Get data...')
            await page.waitForSelector('.dxgv', {visible: true})
            let persons = await page.evaluate(() => {
                document.getElementById("imagelat").style.display = "none"
                let elements = document.querySelectorAll('td[class=dxgv] > table > tbody > tr > td > ul > li')
                
                const dataPersons = []
                for (let element of elements){
                    dataPersons.push(element.innerText)
                }
                return dataPersons
            })
            
            //Structurng data
            console.log('- Structuring data...')
            
            for (let i = 0 ; i < persons.length ; i++){
                
                const LENGTHDIRECTION = 10
                const LENGTHDEPARTAMENT = 14
                const LENGTHCITY = 8
                const LENGTHPHONENUMBER = 9
                
                const indexDirection = persons[i].indexOf('Direccion:') + LENGTHDIRECTION
                const indexDepartament = persons[i].indexOf('Departamento:') + LENGTHDEPARTAMENT
                const indexCity = persons[i].indexOf('Ciudad:') + LENGTHCITY
                const indexphoneNumber = persons[i].indexOf('Telefóno:') + LENGTHPHONENUMBER
                
                const name = persons[i].substr(0 , (indexDirection - LENGTHDIRECTION))
                const direction = persons[i].substr(indexDirection , (indexDepartament -  indexDirection) -  LENGTHDEPARTAMENT)
                const departament = persons[i].substr(indexDepartament , (indexCity-indexDepartament)- LENGTHCITY )
                const city = persons[i].substr(indexCity , ((indexphoneNumber - indexCity ) - LENGTHPHONENUMBER ))
                const phoneNumber = persons[i].substr(indexphoneNumber)
                
                data.push({
                    name,
                    direction,
                    departament,
                    city,
                    phoneNumber,
                })
            }
            
            //save data
            
            const workbook = new Excel.Workbook()
            let worksheet = workbook.addWorksheet('data')
            
            worksheet.columns = [
                {header: 'Name', key: 'name'},
                {header: 'Direction', key: 'direction'},
                {header: 'Departament', key: 'departament'},
                {header: 'City', key: 'city'},
                {header: 'Phone Number', key: 'phoneNumber'},
            ]
            
            //Dump all the data into Excel
            data.forEach((e, index) => {
                //Row 1 is the header
                const rowIndex = index + 2
                
                worksheet.addRow({
                    ...e,
                    amountRemaining: {
                        formula: `=C${rowIndex}-D${rowIndex}`
                    },
                    percentRemaining: {
                        formula: `=E${rowIndex}/C${rowIndex}`
                    }
                })
            })
            
            //Check if workbook is full
            if(data.length > (ROWLENGTHTOSAVE-4)){
                console.log('- Saving data... in workbook ' + workbookNumber)
                await workbook.xlsx.writeFile(`data/Data${workbookNumber}.xlsx`)
                workbookNumber++
                data = []
            }else{
                console.log('- Saving data... in workbook ' + workbookNumber)
                await workbook.xlsx.writeFile(`data/Data${workbookNumber}.xlsx`)
            }
            
            //go to the next
            await page.click('.dxWeb_pNext')
            console.log('- Waiting for the next page...')
            console.groupEnd()
            await page.waitForTimeout(30000)
        }

        console.log('End proceess')
        browser.close()
        
    } catch (error) {
        console.error(error.message)
    }
})()