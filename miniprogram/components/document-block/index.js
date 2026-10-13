Component({properties:{block:Object,issueCount:Number,active:Boolean},methods:{tap(){this.triggerEvent('tap',{blockId:this.data.block.id});}}});
