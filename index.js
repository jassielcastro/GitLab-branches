import fetch from 'node-fetch';
import xl from 'excel4node';

async function fetchBracnhes() {
    console.log("Branches request");
    var page = 1;
    var array = [];
    while (true) {
      var result = await branches(page);
      result.forEach(element => array.push(element));
      if (!result.length || result.length < 100) {
        break;
      } else {
        page++;
      }
    }
    printBranches(array);
  }
  
  async function branches(page) {
    const gitlabToken = 'glpat-1j7rccgJZoAeLxAdyVZk';
    var request = await fetch('https://gitlab.com/api/v4/projects/10654822/repository/branches?per_page=100&page=' + page, {
              method: 'GET',
              headers: {
                  "Authorization": "Bearer " + gitlabToken,
                  "Content-type": 'application/x-www-form-urlencoded'
              }
          });
  
    return await request.json();
  }
  
  function printBranches(branches) {  
    var wb = new xl.Workbook();

    var options = {
        sheetFormat: {
            defaultRowHeight: 25
        }
      };

    var ws = wb.addWorksheet('Branches', options);

    var style = wb.createStyle({
        font: {
            bold: true,
            color: '#000000',
            size: 14,
            family: "roman"
        },
        alignment: {
            horizontal: 'center'
        }
      });

    ws.column(1).setWidth(80);
    ws.column(2).setWidth(40);
    ws.column(3).setWidth(20);
    ws.column(4).setWidth(40);
    
    ws.cell(1, 1).string('Branch name').style(style);
    ws.cell(1, 2).string('Author').style(style);
    ws.cell(1, 3).string('Protected').style(style);
    ws.cell(1, 4).string('Created at').style(style);

    let styleSmall = wb.createStyle({
        font: {
            color: '#000000',
            size: 12,
            family: "roman"
        },
        alignment: {
            horizontal: 'center'
        }
      });

      let styleLarge = wb.createStyle({
        font: {
            bold: true,
            color: '#000000',
            size: 12,
            family: "roman"
        },
        alignment: {
            horizontal: 'left'
        }
      });

      let styleDate = wb.createStyle({
        alignment: {
            horizontal: 'center'
        }, 
        numberFormat: 'dd MMMM yyyy'
      });

    var row = 2;
    sort(filter(branches)).forEach( function(element) {
        ws.cell(row, 1).string(element.name).style(styleLarge);
        ws.cell(row, 2).string(element.commit.author_name).style(styleSmall);
        ws.cell(row, 3).string("" + element.protected).style(styleSmall);
        ws.cell(row, 4).date(Date.parse(element.commit.created_at)).style(styleDate);
        row++;
    });

    wb.write('kavak-branches.xlsx');
}

function filter(branches) {
    return branches
      .filter(branch => !branch.name.includes('release'))
      .filter(branch => branch.merged)
}

function sort(items) {
    return items.sort(function (a, b) {
        let first = a.commit.author_name.toLowerCase()
        let second = b.commit.author_name.toLowerCase()
        if (first > second) {
          return 1;
        }
        if (first < second) {
          return -1;
        }
        return 0;
      });
}
  
fetchBracnhes();
