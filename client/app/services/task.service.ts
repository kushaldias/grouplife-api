import {Injectable} from '@angular/core';
import {Http, Headers, Response} from '@angular/http';
import 'rxjs/add/operator/map';

@Injectable()
export class TaskService{

    log_status = false;

    constructor(private http:Http){
        console.log('task service initialized');
    }
    
    getinfo(info: any){
    
    var headers = new Headers();
    headers.append('Content-Type','application/json');
    
    return this.http.post('http://gpl-api.apps.reactive-solutions.xyz/api/write', info,{headers:headers}).map(res => res.json());
    //return this.http.post('http://172.16.3.69:3000/api/write', info,{headers:headers}).map(res => res.json());
    
    }
    
    
}