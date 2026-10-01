// In-memory implementation of the external FormApp boundary; answers keep item identity.
export class NativeForm {
 constructor(id, titles=[]) { this.id=id;this.items=[];this.next=100;this.description='';this.collects=true;this.quiz=false;this.accepting=true;this.mutations=0;titles.forEach(t=>this.addTextItem().setTitle(t));this.mutations=0; }
 getId(){return this.id;} getTitle(){return this.id;} getDescription(){return this.description;}
 collectsEmail(){return this.collects;} isQuiz(){return this.quiz;}
 getItems(){return [...this.items];} getItemById(id){return this.items.find(i=>i.id===Number(id))||null;}
 getDestinationId(){return 'test-admin-sheet';} isAcceptingResponses(){return this.accepting;} setAcceptingResponses(v){this.accepting=v;return this;}
 setDescription(v){this.description=v;this.changed('description');return this;}
 changed(action,item){this.mutations++;this.onChange?.(action,item);}
 add(type){const i=new NativeItem(this,++this.next,type);this.items.push(i);this.changed('add',i);return i;}
 addTextItem(){return this.add('TEXT');} addParagraphTextItem(){return this.add('PARAGRAPH_TEXT');}
 addMultipleChoiceItem(){return this.add('MULTIPLE_CHOICE');} addListItem(){return this.add('LIST');} addCheckboxItem(){return this.add('CHECKBOX');}
 addGridItem(){return this.add('GRID');} addCheckboxGridItem(){return this.add('CHECKBOX_GRID');} addPageBreakItem(){return this.add('PAGE_BREAK');}
 addSectionHeaderItem(){return this.add('SECTION_HEADER');} addScaleItem(){return this.add('SCALE');}
 addDateItem(){return this.add('DATE');}
 moveItem(item,index){const i=typeof item==='number'?this.items[item]:item;this.items.splice(this.items.indexOf(i),1);this.items.splice(index,0,i);this.changed('move',i);return i;}
 deleteItem(item){const i=typeof item==='number'?this.items[item]:item;this.items.splice(this.items.indexOf(i),1);this.changed('delete',i);return this;}
}
class NativeItem {
 constructor(form,id,type){Object.assign(this,{form,id,type,title:'',help:'',required:false,choices:[],other:false,rows:[],columns:[],bounds:[1,5],labels:['',''],navigation:'CONTINUE'});}
 getId(){return this.id;} getType(){return this.type;} getTitle(){return this.title;} getHelpText(){return this.help;} getIndex(){return this.form.items.indexOf(this);}
 setTitle(v){const old=this.title;this.title=v;this.form.onTitle?.(this,old);this.form.changed('title',this);return this;}
 setHelpText(v){this.help=v;this.form.changed('help',this);return this;}
 isRequired(){return this.required;} setRequired(v){this.required=v;this.form.changed('required',this);return this;}
 includesYear(){return this.year===true;} setIncludesYear(v){this.year=v;this.form.changed('year',this);return this;}
 getChoices(){return this.choices;} createChoice(value,nav){return {getValue:()=>value,getGotoPage:()=>typeof nav==='object'?nav:null,getPageNavigationType:()=>typeof nav==='string'?nav:null};}
 setChoices(v){this.choices=v;this.form.changed('choices',this);return this;}
 hasOtherOption(){return this.other;} showOtherOption(v){this.other=v;this.form.changed('other',this);return this;}
 getRows(){return this.rows;} getColumns(){return this.columns;} setRows(v){this.rows=[...v];this.form.changed('rows',this);return this;} setColumns(v){this.columns=[...v];this.form.changed('columns',this);return this;}
 getLowerBound(){return this.bounds[0];} getUpperBound(){return this.bounds[1];} setBounds(a,b){this.bounds=[a,b];this.form.changed('bounds',this);return this;}
 getLeftLabel(){return this.labels[0];} getRightLabel(){return this.labels[1];} setLabels(a,b){this.labels=[a,b];this.form.changed('labels',this);return this;}
 getPageNavigationType(){return typeof this.navigation==='string'?this.navigation:'GO_TO_PAGE';} getGoToPage(){return typeof this.navigation==='object'?this.navigation:null;} setGoToPage(v){this.navigation=v;this.form.changed('navigation',this);return this;}
 asTextItem(){return this;} asParagraphTextItem(){return this;} asMultipleChoiceItem(){return this;} asListItem(){return this;} asCheckboxItem(){return this;} asGridItem(){return this;} asCheckboxGridItem(){return this;} asPageBreakItem(){return this;} asSectionHeaderItem(){return this;} asScaleItem(){return this;}
 asDateItem(){return this;}
}
