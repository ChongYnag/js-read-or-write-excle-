import React, { Component } from 'react';
import { Button, Icon, message } from 'antd';
import * as XLSX from 'xlsx';

class Excel extends Component {
    onImportExcel = file => {
        // 获取上传的文件对象
        const { files } = file.target;
        // 通过FileReader对象读取文件
        const fileReader = new FileReader();
        fileReader.onload = event => {
            try {
                const { result } = event.target;
                // 以二进制流方式读取得到整份excel表格对象
                const workbook = XLSX.read(result, { type: 'binary' });
                // 存储获取到的数据
                let data = [];
                // 遍历每张工作表进行读取（这里默认只读取第一张表）
                for (const sheet in workbook.Sheets) {
                    // esline-disable-next-line
                    if (workbook.Sheets.hasOwnProperty(sheet)) {
                        // 利用 sheet_to_json 方法将 excel 转成 json 数据
                        data = data.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                        // break; // 如果只取第一张表，就取消注释这行
                    }
                }
                // 最终获取到并且格式化后的 json 数据
                message.success('上传成功！')
                console.log(data);
            } catch (e) {
                // 这里可以抛出文件类型错误不正确的相关提示
                message.error('文件类型不正确！');
            }
        };
        // 以二进制方式打开文件
        fileReader.readAsBinaryString(files[0]);
    }
    //导出excel
    downloadExl(json, type) {
        let outFile = document.getElementById('downlink');
        let keyMap = [] // 获取键
        for (let k in json[0]) {
            keyMap.push(k)
        }
        let tmpdata = [] // 用来保存转换好的json
        json.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
            v: v[k],
            position: (j > 25 ? this.getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
        }))).reduce((prev, next) => prev.concat(next)).forEach(function (v) {
            tmpdata[v.position] = {
                v: v.v
            }
        })
        let outputPos = Object.keys(tmpdata)  // 设置区域,比如表格从A1到D10
        let tmpWB = {
            SheetNames: ['mySheet'], // 保存的表标题
            Sheets: {
                'mySheet': Object.assign({},
                    tmpdata, // 内容
                    {
                        '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] // 设置填充区域
                    })
            }
        }
        let tmpDown = new Blob([this.s2ab(XLSX.write(tmpWB,
            { bookType: (type === undefined ? 'xlsx' : type), bookSST: false, type: 'binary' } // 这里的数据是用来定义导出的格式类型
        ))], {
                type: ''
            })  // 创建二进制对象写入转换好的字节流
        let href = URL.createObjectURL(tmpDown)  // 创建对象超链接
        outFile.download = '文件名.xlsx'  // 下载名称
        outFile.href = href  // 绑定a标签
        outFile.click()  // 模拟点击实现下载
        setTimeout(function () {  // 延时释放
            URL.revokeObjectURL(tmpDown) // 用URL.revokeObjectURL()来释放这个object URL
        }, 100)
    }

    s2ab(s) { // 字符串转字符流
        var buf = new ArrayBuffer(s.length)
        var view = new Uint8Array(buf)
        for (var i = 0; i !== s.length; ++i) {
            view[i] = s.charCodeAt(i) & 0xFF
        }
        return buf
    }

    render() {
        let json = [{ "充值编号": "880ccbd5ca40a0d06b01dd8449de2e4c", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 11:12:43", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "4f50477420f4bf5793d5a588a760882f", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 11:08:30", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "b7f3eac68413e8cf654c93bf251e6953", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 11:04:06", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "9589abd17018cdab2a3151e74f582577", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:59:42", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "6350f04eb3a76d1ecc7b362b8f76dcfa", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:55:29", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "e68ae1f6fd99876be92377130ad8ad6a", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:51:05", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "60081fb6175306be907c5eb660aee628", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:46:41", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "3559fc02c03fc11fe514870efaeaf593", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:42:17", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "1c0073712772dd2a7c6a2772ccd000d9", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:37:53", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "34f636f9663762062322d7a42926c78b", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:33:29", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "408ab56bf2767f9018515e07dae70b4d", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:29:16", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "008493c15042185932353a880a59fc03", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:24:52", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "388296c7a62a471f40bc85ebcad1569f", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:20:28", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "65146f1eccb45f76cfec9ffb230d15e2", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:16:15", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "354265620285cc03ba40953113bde94b", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:11:51", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "34a0718470eddc7b5ccd740efa49268e", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:07:27", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "5c11af0a0064cdd6b84248426c46d13f", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 10:03:04", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "8bfd0328b273cdc0ac09fea7b8621f17", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:58:50", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "cb96b38544cd0b4cc676eaec5aea17b5", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:54:26", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "0b9531e4015e6cad85940305d728d0d5", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:50:02", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "71d5735de1f765277741ff73372bd797", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:45:49", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "32d36d7a05853501ffe36de1ae0d8cec", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:41:25", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "cb15fd5f7d2e3ab2dddad89e7f847e6b", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:37:01", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "98b69a73461637db157884c6de8197a1", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:32:48", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "c7bc3123c6771ee15e30122ed15caa44", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:28:24", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "14cb854b738321a4b8aa0247816b05e5", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:24:00", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "c92920516bd1a8127668e95ea170556c", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:19:47", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "6dbe1757273670f9496828024cc239e3", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:15:23", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "b18ee4fb6f78ea9abf5c8b53366cfccf", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:10:59", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "d502c8efbc3044c7dcfe66e1294f8ffc", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:06:46", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "9247e160a7ad53dd8842c3acf16d0755", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 09:02:22", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "ac6d523593bbaa5deb0728215bb3b95b", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:57:58", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "52e13c0e07f9fd5c3a7e4d4fe6d36623", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:53:45", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "e9bd9bc678a367c0d940c410372a7f1a", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:49:21", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "5a5e6aeb22e7b367003ded71191ee96e", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:44:57", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "eb459fbc4e8ee09f2fc6761f15dfdf1a", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:40:33", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "a2c162f54f431e0c8e71ba04da60a1d8", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:36:09", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "d1a5ade6ab538ddac383647cdea8f6d4", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:31:45", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "b590c275b486afd81c59d5c75ad49e49", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:27:21", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "3637cae4760b6b888fb6e8186e3e039a", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:22:57", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }, { "充值编号": "8c41c03146344087b0180ab78b7f923a", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 08:18:33", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" }];
        for (let i = 0; i < 10000; i++) {
            json.push({ "充值编号": "880ccbd5ca40a0d06b01dd8449de2e4c", "uid": 58, "用户邮箱": "zhangkewei@onething.net", "用户手机": "86|15210923819", "用户名": "zhangkewei@onething.net", "充值地址(from地址)": 0, "提交时间": "2019-02-26 11:12:43", "充值数量(CJF)": 0, "支付状态": "确认异常，不需要处理" })
        }
        return (
            <div style={{ marginTop: 100 }}>
                <Button>
                    <Icon type='upload' />
                    <input type='file' accept='.xlsx, .xls' onChange={this.onImportExcel} />
                    <span>上传文件</span>
                </Button>
                <p>支持 .xlsx、.xls 格式的文件</p>
                <Button onClick={() => this.downloadExl(json)}>导出</Button>
                <a href="javascript:" id="downlink"></a>
            </div >
        );
    }
}

export default Excel;
