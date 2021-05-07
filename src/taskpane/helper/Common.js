export default function ConvertNumberToCurrenty(n){
    //$ {(Math.round(item.fields.Estimated_x0020_Project_x0020_Va * 100) / 100).toLocaleString()}
    if(n){
        const numb=(Math.round(n * 100) / 100).toLocaleString();
        return "$ " + numb+"";
    }
    else{
        return "$0";
    }
}