Component({properties:{issue:Object},methods:{accept(){this.triggerEvent('accept',{issueId:this.data.issue.id});},reject(){this.triggerEvent('reject',{issueId:this.data.issue.id});}}});
