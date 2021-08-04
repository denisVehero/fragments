export class ProgressStatus {
	planed: number;
	complited: number;
	title: string;
	constructor (planed: number, complited: number, title: string) {
		this.planed = planed;
		this.complited = complited;
		this.title = title;
	}
}
