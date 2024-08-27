const {exec} = require('child_process')
const {PowerShell} = require('node-powershell')
var fs = require('fs')
var bodyParser = require('body-parser')
var express = require('express')

const app = express()
app.use(express.json())
app.listen(
    7777,
    '0.0.0.0'
)

var config = {}
fs.readFileSync('../config.txt', 'utf-16le').toString().split(/\r?\n/).forEach(line => {
    if(line.length==0 || line.startsWith('#') || line.startsWith('//'))
        return
    arr = line.split('=')
    config[arr[0]]=arr[1]
})
console.log(config)

const ps = new PowerShell({
    executionPolicy: 'Bypass',
    noProfile: true
})
const hymnsPath = config['hymnLyricsPath']

app.get('/', (req,res) => {
    console.log("test request recieved")
    res.status(200).json({
        "success":true
    })
})

app.post('/change_slide', async (req,res) => {
    console.log(req['body'])
    
    const command = PowerShell.command`& "../PPT scripts/Change-All-PPT-Slides.ps1" -count ${req['body']['count']}`
    const output = await ps.invoke(command)
    console.log(output)
    console.log(output['stdout'].toString())
    console.log(output['stderr'].toString())
    // exec('& "../PPT scripts/Change-All-PPT-Slides.ps1" -count 1', {'shell':'powershell.exe'},(error,stdout,stderr) => {
    //     console.log(error)
    //     console.log(stdout)
    //     console.log(stderr)
    // })
    res.send("reponse")
})

app.get('/hymns', (req,res) => {
    console.log("request recieved")
    res.status(200).json({
        hymns : fs.readdir(hymnsPath, (err, files) => {
            if(!err){
                res.status(200).json({
                    hymns : files
                })
            }
        })
    })
})
