/* 
Objetos
facilitação de acesso a dados 
separação especifica de dados para melhor controle 



*/

const name = "Arthur"
const age  = 19
const adress = "rua alcindo nardini 189"



const arthur = { 

    name: "Arthur",
    age: 19,
    adress: {
        street: "rua alcindo nardini",
        number: 189
    }
}
arthur.age = 16
console.log(arthur)