import { Component } from '@angular/core';
import '../../public/css/styles.css';
import {AuthorizationService} from './authorization.service.ts';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  providers: [AuthorizationService]
})
export class AppComponent { 

	private armResourceId: string = "797f4846-ba00-4fd7-ba43-dac1f8f63013";

	constructor(private authorizationService: AuthorizationService){}

	authorize() {
		this.authorizationService.signIn(this.userSignedIn);
	}

	private userSignedIn(err:string, token:any){
		console.log(token);
	}

	getCachedUser(){
		console.log(this.authorizationService.getCachedUser());
	}

	getARMAccessToken(){
		this.authorizationService.acquireToken(this.accessTokenCallback, this.armResourceId);
	}
	
	private accessTokenCallback(err:string, token:any){
		console.log(token);
	}


}
