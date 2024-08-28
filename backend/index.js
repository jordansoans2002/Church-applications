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

const ps = new PowerShell({
    executionPolicy: 'Bypass',
    noProfile: true
})

app.get('/', (req,res) => {
    console.log("test request recieved")
    res.status(200).json({
        "success":true
    })
})

app.post('/change_slide', async (req,res) => {
    console.log(req['body'])
    
    const count = req.body.count
    const option = req.body.option
    const command = PowerShell.command`& "../PPT scripts/Change-All-PPT-Slides.ps1" -count ${count} -pptOption ${option}`
    const output = await ps.invoke(command)
    const stdout = output.stdout.toString()
    console.log(stdout)
    const statusCode = Number(stdout.substring(0,stdout.indexOf(':')))
    const msg = stdout.substring(stdout.indexOf(':')+1)
    res.status(statusCode).json(
        {
            "message":msg
        }
    )

    // console.log(stdout.toString('utf16le'))
    // const dict = JSON.parse(stdout.toString('utf8').replace('\n',' '))
    // console.log("stdout:"+typeof(dict)+" "+dict)
    // console.log(typeof(output.stderr.toString()))

    // const responseStr = stdout[stdout.length-1]
    // const statusCode = Number(responseStr.substring(0,responseStr.indexOf(':')))
    // const msg = responseStr.substring(responseStr.indexOf(':')+1)
    // res.status(statusCode).json(
    //     {
    //         "message":msg,
    //         "lyricsLang1":stdout[0],
    //         "lyricsLang2":stdout[1]
    //     }
    // )


    // exec('& "../PPT scripts/Change-All-PPT-Slides.ps1" -count 1', {'shell':'powershell.exe'},(error,stdout,stderr) => {
    //     console.log(error)
    //     console.log(stdout)
    //     console.log(stderr)
    // })
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
