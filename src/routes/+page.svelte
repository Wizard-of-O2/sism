<script lang="ts">
	import * as XLSX from 'xlsx';
	import { BaseDirectory, writeBinaryFile } from '@tauri-apps/api/fs';

	let files: FileList;
	let log = '';
	let tableHtml = '';

	async function handleClick() {
		if (files?.length) {
			const data = await files[0].arrayBuffer();
			const wb = XLSX.read(data);
			const ws = wb.Sheets[wb.SheetNames[0]];
			tableHtml = XLSX.utils.sheet_to_html(ws, {
				id: 'my-table',
				header: undefined,
				footer: undefined
			});
		}
	}

	async function handleSave() {
		if (tableHtml) {
			log = 'Saving...';
			const tableElem = document.getElementById('my-table');
			const wb = XLSX.utils.table_to_book(tableElem);

			const bytes = XLSX.writeXLSX(wb, {
				type: 'array',
				compression: true
			});

			await writeBinaryFile('a.xlsx', bytes, {
				dir: BaseDirectory.Download
			});
			log = 'Finish!';
		}
	}
</script>

<h1>An Excel App</h1>
<div>
	<label for="file">대상 파일</label>
	<input type="file" id="file" accept=".xls,.xlsx" bind:files />
</div>
<div>
	<p>이 버튼을 눌러 엑셀 파일을 로드하세요</p>
	<button on:click={handleClick}>처리</button>
</div>
<div>
	<p>이 버튼을 누르면 로딩된 엑셀을 다운로드 폴더에 a.xlsx로 저장합니다.</p>
	<button on:click={handleSave}>저장</button>
</div>
<div>
	{#if tableHtml}
		{@html tableHtml}
	{/if}
</div>
<div>
	<p>{log}</p>
</div>
