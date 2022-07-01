const request = require('request');
const jsdom = require('jsdom');
const { JSDOM } = require('jsdom');
const ExcelJS = require('exceljs');

//  Exel init
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Articles');
sheet.columns = [
    { header: 'Tên báo', key: 'article_name', width: 10},
    { header: 'Số báo', key: 'article_number', width: 10},
    { header: 'Ngày đăng', key: 'date_published', width: 10},
    { header: 'DOI', key: 'doi', width: 10}
];

const url = 'https://jprp.vn/index.php/JPRP/issue/archive';
request(url, (err, res, body) => {
    let mainPage = body;
    // console.log(mainPage);

    const mainPageDom = new JSDOM(mainPage);
    console.log(mainPageDom.window.Document);

    //  Link tất cả các tập
    let episodeLinkArray = mainPageDom.window.document.querySelectorAll('.media-body a');

    //  Vào từng tập
    for (let i = 0; i < episodeLinkArray.length; ++i) {
        
        request(episodeLinkArray[i].href, (err, res, body) => {
            console.log(episodeLinkArray[i].href);
            const episodePageDom = new JSDOM(body);
            let allNewsLink = episodePageDom.window.document.querySelectorAll('.media-body .row .col-md-10 a');
            console.log(allNewsLink.length);
            for (let j  = 0; j < allNewsLink.length; ++j) {
                
                request(allNewsLink[j].href, (err, res, body) => {
                    console.log(allNewsLink[j].href);
                    const newsDom = new JSDOM(body);
                    let title = '';
                    let datePublished = '';
                    let articleNumber = '';
                    let doi = '';
                    try {
                        title = newsDom.window.document.querySelector('.article-details header h2').innerHTML.trim();
                    } catch (error) {}
                    try {
                        datePublished = newsDom.window.document.querySelector('.list-group-item.date-published').innerHTML.trim();
                        datePublished = datePublished.substring(datePublished.lastIndexOf('</strong>') + 9, datePublished.length);
                        datePublished = datePublished.trim();
                    } catch (error) {}
                    try {
                        articleNumber = newsDom.window.document.querySelector('.issue .panel-body a').innerHTML.trim();
                    } catch (error) {}
                    try {
                        doi = newsDom.window.document.querySelector('.list-group-item.doi a').innerHTML.trim();
                    } catch (error) {}
                    
                    console.log(title);
                    console.log(datePublished);
                    console.log(articleNumber);
                    console.log(doi);
                    sheet.addRow({
                        article_name: title,
                        article_number: articleNumber,
                        date_published: datePublished,
                        doi: doi
                    });
                    
                    workbook.xlsx.writeFile('index.xlsx');
                });
                // break;
            }
        });
        // break;
    }
});