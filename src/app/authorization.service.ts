import { Injectable } from '@angular/core';
import {remote} from 'electron';

declare var AuthenticationContext:any; //Make an existing javascript library available in TS... no net

@Injectable()
export class AuthorizationService {

	private authContext: any;
	private authWindow: Electron.BrowserWindow;

	public constructor() {
		this.authContext = new AuthenticationContext( {
			instance: 'https://login.microsoftonline.com/',
      		tenant: 'common', //COMMON OR YOUR TENANT ID
      		clientId: '188743b4-4d95-4971-9d7b-1f2c105c9bda', //REPLACE WITH YOUR CLIENT ID
      		redirectUri: 'http://localhost', //REPLACE WITH YOUR REDIRECT URL
      		displayCall: this.electronSignIn.bind(this),
      		callback: this.userSignedIn.bind(this)
		})

		var webPreferences: Electron.WebPreferences = {};
		var options: Electron.BrowserWindowOptions = {};
		options.width = 800;
		options.height = 600;
		options.show = false;
		webPreferences.nodeIntegration = false;
		options.webPreferences = webPreferences;

		this.authWindow = new remote.BrowserWindow(options);

		var handleRedirect = function(event: Electron.Event, oldURL:string, newURL:string, isMainFrame:boolean, httpResponseCode:number, requestMethod:string, referrer:string, headers:Electron.Headers){

			this.handleCallBack(newURL);
			console.log('redirect detected');
		};

		var handleNavigation = function(event: Electron.Event, url:string, that:any){
			this.handleCallBack(url);
			console.log('navigation detected');

		};

		var handleRedirectBound = handleRedirect.bind(this);
		var handleNavigationBound = handleNavigation.bind(this);
		
		this.authWindow.webContents.on("did-get-redirect-request", handleRedirectBound);

		this.authWindow.webContents.on("will-navigate", handleNavigationBound);
		
	}


	private handleCallBack(url:string){
		if (url.indexOf(this.authContext.config.redirectUri) != -1) {

			this.authContext.handleWindowCallback(url.substr(url.indexOf('#')));
			this.authWindow.hide();



		}
	}

	private electronSignIn(url:string){
		console.log(url);

		this.authWindow.loadURL(url);
		this.authWindow.show();

	}

	private userSignedIn(user:any){
		console.log(user);
	}

	public signIn(){
		this.authContext.login();
	}

	public acquireToken(resourceId: string){

	}


}