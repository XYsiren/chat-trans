Component({properties:{block:Object,issueCount:Number},methods:{tap(){this.triggerEvent('tap',{blockId:this.data.block.id});}}});
