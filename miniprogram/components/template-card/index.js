Component({properties:{template:Object,selected:Boolean},methods:{choose(){this.triggerEvent('choose',{id:this.data.template.id});}}});
