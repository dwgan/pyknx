import asyncio
from xknx import XKNX
from xknx.io import ConnectionConfig, ConnectionType
from xknx.dpt import DPTBinary
from xknx.telegram import GroupAddress, Telegram
from xknx.telegram.apci import GroupValueWrite


async def send_knx_command():
    # 配置连接参数 (匹配您的抓包信息)
    connection_config = ConnectionConfig(
        gateway_ip="192.168.0.11",  # KNX路由器IP
        gateway_port=3671,  # KNX端口
    )

    # 创建XKNX实例
    async with XKNX(connection_config=connection_config) as xknx:
        # 创建目标组地址
        group_address = GroupAddress("5/1/1")

        # 创建有效载荷 (关灯命令$00)
        payload = GroupValueWrite(value=DPTBinary(1))

        # 创建并发送Telegram
        telegram = Telegram(
            destination_address=group_address,
            payload=payload
        )
        await xknx.telegrams.put(telegram)
        print("✅ 命令已成功发送到KNX总线")


# 运行异步函数
asyncio.run(send_knx_command())