declare module '*.jpg';
declare module '*.jpeg';
declare module '*.png';
declare module '*.svg';
declare module '*.gif';
// declare module '*.mp4';
declare module '*.scss' {
	const styles: { [className: string]: string };
	export default styles;
}
declare module '*.module.scss' {
	const styles: { [className: string]: string };
	export default styles;
}