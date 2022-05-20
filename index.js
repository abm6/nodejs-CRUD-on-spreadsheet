const xlsx = require('xlsx');
const fs = require('fs');

const fileName = 'Book1.xlsx';
var outputFileName = 'Output.xlsx';

const file = xlsx.readFile(`./${fileName}`);

var data = [];

const sheets = file.SheetNames;

const fetchData = async () => {
	for (let i = 0; i < sheets.length; i++) {
		const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
		temp.forEach((res) => {
			data.push(res);
		});
	}
};

function updateColumn(arr = []) {
	const id = parseInt(arr[0]);
	const key = arr[1];
	const value =( key == 'Score' || key == 'Id' )? parseInt(arr[2]) : arr[2];

	if (Object.keys(data[0]).includes(key)) {
		data = data.map((item) => {
			if (item.Id === parseInt(id)) {
				item[key] = value;
			}
			return item;
		});
	} else {
		throw new Error('Invalid column name');
	}
}

function createRow(arr = []) {
	const row = {
		Id: parseInt(arr[0]),
		Name: arr[1],
		Stream: arr[2],
		Track: arr[3],
		Score: parseInt(arr[4]),
	};

	if (isNaN(row.Id)) throw new Error('Invalid Id');
	if (isNaN(row.Score)) throw new Error('Invalid Score');

	try {
		data.find((item) => {
			if (item.Id === row.Id) throw new Error('Id already exists');
		});
		data.push({ ...row });
	} catch (error) {
		console.log(error);
	}
}

function deleteRow(id) {
	data = data.filter((item) => {
		return item.Id !== parseInt(id);
	});
}

function writeToFile(fileName = 'Output.xlsx') {
	const newWB = xlsx.utils.book_new();
	const newWs = xlsx.utils.json_to_sheet(data);

	xlsx.utils.book_append_sheet(newWB, newWs, 'Sheet1');

	// Writing to a separate file so that the original file is not affected
	xlsx.writeFile(newWB, `./${fileName}`);
}

function persist() {
	if (fs.existsSync(`./${outputFileName}`)) {
		fs.unlinkSync(`./${outputFileName}`);
	}
	outputFileName = fileName;
}

async function performCrudOperations(args) {
	const operations = {
		'-c': createRow,
		'-r': () => {},
		'-u': updateColumn,
		'-d': deleteRow,
	};

	if (args[0] in operations) {
		operations[args[0]](args.slice(1));
		if (args[args.length - 1] === '--persist') persist();
	} else if (args[0] == '--persist') {
		persist();
	} else if (args[0] == null || args[0] == '') {
		throw new Error('No arguments provided');
	} else {
		throw new Error('Invalid argument');
	}
}

(function main() {
	const args = process.argv.slice(2);

	fetchData()
		.then(performCrudOperations(args).then(() => console.table(data)))
		.then(writeToFile(outputFileName))
		.catch((e) => console.log(e));
})();
