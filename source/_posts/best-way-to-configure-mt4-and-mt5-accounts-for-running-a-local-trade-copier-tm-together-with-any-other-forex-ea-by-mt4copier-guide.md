---
title: Best way to Configure MT4 and MT5 Accounts for Running a Local Trade Copier™ Together With Any Other Forex EA
date: 2024-05-19T17:47:00.110Z
tags: 
  - mt5
  - mt4
categories: 
  - apps
description: Best way to configure MT4 and MT5 accounts for running a Local Trade Copier™ together with any other Forex EA. In this video, Rimantas explains how to do it. - Best way to Configure MT4 and MT5 Accounts for Running a Local Trade Copier™ Together With Any Other Forex EA
keywords: mt4 to mt5 trade,mt4 to mt5,mt4 to mt5 trade copier,meta trader
---

The [Local Trade Copier](https://tools.techidaily.com/mt4copier/) software is a powerful tool that allows you to copy trades between multiple MetaTrader 4 and MetaTrader 5 accounts. It is a perfect solution for money managers and signal providers who need to manage multiple accounts at the same time. The LTC software is also a great tool for traders who want to copy trades between their own trading accounts.

In this guide, we will show you how to configure your MetaTrader 4 and MetaTrader 5 accounts to run the [Local Trade Copier](https://tools.techidaily.com/mt4copier/) together with any other Forex EA.

<iframe width="898" height="503" src="https://www.youtube.com/embed/TiaPtSBhguQ" title="Set up MT4 &amp; MT5 Accounts to Copy and Paste Trades From One Forex EA Across Many Metatrader Accounts" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

## Pricing

- **Local Trade Copier™ for MT4 & MT5 (Personal Monthly Plan)** : [$31.47/month](https://secure.2checkout.com/order/cart.php?PRODS=4722887&QTY=1&AFFILIATE=108875)
- **Local Trade Copier™ for MT4 & MT5 (Manager Monthly Plan)** : [$96.58/month](https://secure.2checkout.com/order/cart.php?PRODS=4723269&QTY=1&AFFILIATE=108875)
- **Local Trade Copier™ for MT4 & MT5 (VIP Monthly Plan)** : [$215.95/month](https://secure.2checkout.com/order/checkout.php?PRODS=4723270&QTY=1&AFFILIATE=108875)
- **Local Trade Copier™ for MT4 & MT5 (Personal Annual Plan)** : [$314.71/year](https://secure.2checkout.com/order/cart.php?PRODS=4723646&QTY=1&AFFILIATE=108875)
- **Local Trade Copier™ for MT4 & MT5 (Manager Annual Plan)** : [$965.83/year](https://secure.2checkout.com/order/cart.php?PRODS=4723648&QTY=1&AFFILIATE=108875)
- **Local Trade Copier™ for MT4 & MT5 (VIP Annual Plan)** : [$2159.55/year](https://secure.2checkout.com/order/cart.php?PRODS=4723650&QTY=1&AFFILIATE=108875)




## MAAB Trade Filter: Copy Master Account Only When It Is Making Profits

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/1.png)

All strategies have drawdowns or periods when they face unfavorable market conditions. MAAB Trade Filter is specifically made to reduce negative effects on your account when combined with the [Local Trade Copier](https://tools.techidaily.com/mt4copier/). It is an easy solution to minimize drawdown and copy Master Account only when it is making profits. This tool also has the power to turn bad EAs and strategies into winners. Find out how MAAB Trade Filter makes it happen in this guide.

### MAAB Trade Filter – How it Works

MAAB stands for Moving Average on Account Balance. As traders, we open and close trades, and the account balance moves up and down, creating a histogram. Now, plotting a Moving Average on top of the balance histogram we instruct MAAB Trade Filter not to allow losing trades to take effect. Losing trades that make the balance or equity below this Moving Average are simply filtered out. They are not copied from the Master account to the Client account. Once winning trades makes the account balance histogram go above the Moving Average, MAAB Trade Filter will allow them to grow the Client account. A simple, yet unique solution to effectively protect Client accounts.

In the picture below you can see the Master account where the purple background is. MAAB Trade Filter indicator is also visible here. The Client account is just below. This example shows two MAAB Trade Filters applied from two Master accounts.  It is just a demonstration that you can merge multiple MAAB Trade Filters into one Client account.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/2.png)

MAAB Trade Filter tracks one Master account. However, you can stack up multiple MAAB Trade Filters to track multiple Master Accounts. Very useful if you have different strategies from multiple Master accounts and want to filter those who are not currently making profits. This way, the Client EA only receives trades from winning Master accounts. MAAB Trade Filter can also work in reverse mode if you need it.

Note that the red Moving Average you see on the Master account histogram is just for your reference. The Moving Average on the Client-side is the one that will be used by Client EA for filtering. Just to avoid confusion if you use different Moving Average settings on those accounts.

### Setting up your MAAB Trade Filter

1. First things first, set your MAAB Trade Filter indicator on the server-side, on your Master account. The MAAB Trade Filter will scan every closed trade on the Master account and then send the signal to Client accounts, if applicable.
2. Secondly, plug-in MAAB Trade Filter to your Client account too. MAAB Trade Filter histogram you see on the Client account shows the account balance from the Master account, not the Client account balance. The reason behind this is to show you what trades are filtered out from the Master account.
3. After you plug in MAAB Trade Filter on the Client-side you need to type in the account number you want it to track. In the top right part of the picture below, we use account number 60055865 as an example:
   ![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/3.png)
4. Now, we need to enable MAAB Trade Filter on the Client EA. It may look like it is automatically enabled once you see the visuals, but we need to turn it on from the Client EA settings window. Scroll to the Trade Filter section and set it to True.
   
We are ready now to put MAAB Trade Filter into action! But make sure you understand all of MAAB Trade Filter powers and how to adjust it to your preference. In the next section, we will demonstrate exactly that.


### MAAB Trade Filter Features Demonstration

So, I have my VPS up and four MetaTrader accounts running on it. MetaTrader 4 will be my Master account #1, and MetaTrader 5 will be my Master account #2. On the right side, I have two Client accounts, running on MetaTrader 4 and MetaTrader 5.  Now, both Client accounts will copy trades from both Master accounts, however, MAAB Trade Filter will be engaged so we can see how it filters undesirable trades.

First, I will set up my server-side Master accounts. Open MetaTrader 4 and apply Server EA to the chart (drag n drop) from the Experts list. You will need to have at least version 2.9.9 of the EA to make it work with the MAAB Trade Filter. Click OK and do not worry about settings now.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/4.png)


Now go to the MT4 Indicator list and apply the “MA on Account Balance (server)” indicator to the chart too. No need to change the settings, however, if you need it to track an EA with a specific Magic Number you have that option available. Otherwise, a setting of 0 means it tracks manual trading. If you set the Magic Number input to -1 it will track all the trades on this account.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/5.png)

If you zoom into the histogram, you will notice each histogram bar represents how closed trade affected the balance. A mouse-over tooltip above the indicator will display balance information on each bar. Notice that the histogram went down as losing trades closed and it went just below the Moving Average. At that point, MAAB Trade Filter would stop accepting trade signals from this Master account.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/6.png)

Now I will set up a Metatrader 5 account, my second Master account. I can repeat the same procedure as for the first Master account. I apply the Server EA for the MT5 (at least version 1.1.6), then attach the MAAB Trade Filter indicator dedicated for the MT5 server-side. Notice this second Master account has a balance histogram below the filtering Moving Average, meaning it has a series of losing trades. We do not want these losing trades on the Client account, don’t we? That is why we filter them out using the MAAB Trade Filter.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/7.png)

Moving on to the Client account on MT5. Here I will first insert the MAAB Trade Filter indicator for the client-side onto the MT5 chart window. A settings window will pop up and it will show a few options. You can change the Moving Average periods (default 13), MA types like Simple, Exponential, Smoothed, etc. For the MAAB Trade Filter to work, we must input the Master account number in the ServerAccountNumber field. No worries, in case you forget to type in the number the indicator window will display a warning.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/8.png)

In my example, the Master account number is 60055865. Again, the indicator will now show the Master account balance, not the Client account balance.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/9.png)

Finally, I can attach the Client EA now onto the chart. In the EA settings, scroll down and find the MA Trade Filter line and set it to True. Right below you will also see the option to apply the Moving Average to the Master account balance histogram or account’s equity.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/10.png)

Regardless of your preference, The MAAB Trade Filter indicator will show an orange horizontal line that represents the current equity of the Master account. It will refresh every 15 seconds or so. As we see, the balance of this Master account is above the Moving Average so the Client EA will copy the trades to the client side.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/11.png)

If you want to set the Moving Average to 50 periods, of course, this will also affect how the MAAB Trade Filter indicator behaves. With a 50-period Moving Average, the histogram balance is now below the MA, meaning the MAAB Trade Filter will cease copying trades from this Master account. Since the 50 period MA reacts slower to the histogram changes, it will need more winning trades before the histogram is above the MA(50). Only then the Master account trades will be allowed again to the Client account.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/12.png)

For now, let’s change back the MA period settings to 13.

While I am still on this client-side platform, I will add the MAAB Trade Filter for the second Master account (the MT5). I will insert the MAAB indicator (client-side) into the chart and input the Master account number in the settings.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/13.png)

Now we can see two indicator windows showing the balance histogram from each Master account.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/14.png)

When we have a Client account setup like this, the MAAB Trade Filters will cease copying trades for any Master account that does not qualify. To qualify they need to have a histogram above the Moving Average you set.

Finally, we have a second Client account that we want to improve with MAAB Trading Filter. I will attach the MAAB Trade Filter (client-side) indicator first and then set it to connect with the first #60055865 Master account in the ServerAccountNumber field.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/15.png)

Now we add the Client EA v2.9.9f from the Experts list and enable it from the settings window. This time I will also set the EA to compare the Moving Average to equity instead of the account balance.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/16.png)

Check out the orange Equity line. It is below the Moving Average right?

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/17.png)

Because of this, all trades from this Master account will be suspended until the equity goes above the Moving Average.

### MAAB Trade Filter in Action Examples

Let’s go ahead and make some trades to see how MAAB Trade Filter manages trading from winning Master accounts and from those that currently do not show good performance. I already have a lot of trades open so I will pick one currency pair that does not have any. It is the USDJPY.

Let’s buy half a lot on the first Master account (#60055865) and see what happens on the Client-side platforms.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/18.png)

As expected, the first Client account immediately copied the trade. The account balance from the Master account was above the Moving Average.

However, the second Client account denied that trade. This is because we set the rule to compare the MA to Master’s equity – which was below the MA. We can confirm this by looking at the Experts tab and the line that says “Ignored trade BUY USDJPY because of the MAAB Trade Filter. Master account equity below MA13”. Clearly said, it is doing what we set it to do.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/19.png)

Even though the balance histogram was above the MA for this Master account, the equity was not. Once the equity goes above MA(13) the MAAB Trade Filter will allow it to pass to the Client account.

Trades that are ignored by the system will also trigger the question mark on the main chart screen to turn red.  A counter for all ignored trades is available too.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/20.png)

Let’s do some more trading from the other Master account. EURUSD looks like a good example, buying 1 lot. This trade is sent to two Client accounts we have linked, but let’s see if it is filtered on any.

The first Client account ignored this EURUSD trade, but the second Client account accepted it. Simply because we did not set up the second MAAB Trade Filter for this Master MT5 account on the second Client account. All trades that come from the Admiral Markets Master account to the IC Markets Client account are copied without any filtering.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/21.png)

If we go to fullscreen we can see the message in the Experts tab that trades from one of the Master accounts are ignored. The reason is “Master account balance is below MA13”. Perfect! That is what we want.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/22.png)

Now let’s see what happens when we close this EURUSD trade. It is in a small profit. Alright, the histogram went up a bit as the balance increased by the profit amount. The equity was updated too. On the Client side, the same balance and equity changes are also visible.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/23.png)

I will find and close some trades in a loss so I can show you what happens in this case. In the screenshot below, you can see that the loss caused the balance to go down which can be seen on the balance histogram going below the Moving Average.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/24.png)

This means MAAB Trade Filter will no longer allow trades to this Client from either Master account. They both do not pass the filter rules we have set in the MAAB Trade Filter. To test this I will open a new USDJPY trade. As expected I see a “sell USDJPY ignored” message on the Client (see screenshot below), good job!

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/25.png)

Since I have some trades in profit too I will close them to see what happens.

After I close a profitable trade, the balance goes up and its histogram goes above Moving Average again!

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/26.png)

Does it mean the MAAB Trade Filter will now allow this Master account trades to the Client-side? Let’s test it out, I will make another USDJPY trade.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/27.png)

Alright, the first client copied this trade, however, the second Client account did not. If you remember, there is a rule for MAAB Trade Filter we set – If equity is below the MA, ignore trades from this Master account. It just does what is supposed to do.

### MAAB Trade Filter Reverse Logic

I open the Client EA settings on the first Client account (the one with two MAAB Trade Filters) and set the MAAB Reverse Logic parameter to True.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/28.png)

I will also set the Reverse Trades to True under the Trades Manipulation settings section in the Client EA.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/29.png)

Click OK and then open the Client EA settings on the second Client account (the one with just one MAAB Trade Filter set to compare equity). Now, let’s keep the equity rule and change MAAB Reverse Logic to True.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/30.png)

Ok, now the first Client account has the EA instructed to use MAAB Reverse Logic and also reverse trades, meaning “buy” trades become “sell” trades and “sell” trades become “buy” trades.

The second Client account has MAAB Trade Filter set to compare MA to Master account equity but with Reverse Logic – meaning if equity is below the MA trades are copied to the Client account and if above they are filtered. Let’s test this setup with a trade now.

Opening USDJPY “buy” trade…

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/31.png)

The second Client account with the reversed equity rule was allowed to copy trades. Makes sense since the equity is below the MA. Before the Reverse Logic change, all trades were filtered.

Meanwhile, the first Client account ignored this USDJPY trade. The Reverse Logic denied it now since the balance is above the MA. Normally, the MAAB Trade Filter would allow it, however, we have reversed its logic.

But what will happen if I open a EURUSD “buy” trade on the second Master account, the one with reversed trade logic? My “buy” EURUSD trade gets copied as “sell” EURUSD.

![](https://tools.techidaily.com/images/apps/mt4copier/maab-trade-filter/32.png)

The second Client account copied the trade normally since we do not have any MAAB Trade Filter enabled for this Master account.

The first Client account is a different story though. The Master account has a balance histogram below the MA. Normally, it would not be allowed to the Client-side, yet we have inversed its logic. With the Reversed Trade enabled too, the “buy” EURUSD becomes “sell” EURUSD.

After all this, you might be wondering why would we need all this inversion? Well, this is a perfect solution for all Reversal trading strategies! It makes sense to reverse trades from a losing account, right? Why not make bad, losing strategies profitable? And with the MAAB Filter reverse logic, we are going one step further – we copy bad losing strategies (and reverse their trades) only when the Master account is losing. It is funny but true, now you can start composing bad strategies as well, and make use of all those EAs or strategies you have thrown into the trash bin.



## How to set different lot sizes for each copier account

<iframe width="898" height="503" src="https://www.youtube.com/embed/K6JwObVWivU" title="How to set different lot sizes for each copier account" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

In this video, Rimantas explains how you can easily set different lot sizes for each copier account with the [Local Trade Copier](https://tools.techidaily.com/mt4copier/)™ for MT4 & MT5.

## How To Configure MT4 and MT5 Accounts for Running a Local Trade Copier™ Together With Any Other Forex EA

The [Local Trade Copier](https://tools.techidaily.com/mt4copier/) software is a powerful tool that allows you to copy trades between multiple MetaTrader 4 and MetaTrader 5 accounts. It is a perfect solution for money managers and signal providers who need to manage multiple accounts at the same time. The LTC software is also a great tool for traders who want to copy trades between their own trading accounts.

In this guide, we will show you how to configure your MetaTrader 4 and MetaTrader 5 accounts to run the [Local Trade Copier](https://tools.techidaily.com/mt4copier/) together with any other Forex EA.

<iframe width="898" height="503" src="https://www.youtube.com/embed/TiaPtSBhguQ" title="Set up MT4 &amp; MT5 Accounts to Copy and Paste Trades From One Forex EA Across Many Metatrader Accounts" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>


Today, I’m diving into a game-changing tool that’s going to make your trading life a whole lot easier and way more profitable. In this video, I talk about the [Local Trade Copier](https://tools.techidaily.com/mt4copier/)™ and how it’s going to save you from buying a Forex EA license for every single one of your MT4/MT5 accounts. Yes, you read that right!


Imagine this: You’ve got a Forex bot that’s making some really smart trades in Forex, Gold, Oil, and other markets. It’s working its magic on your MT4/MT5 account, and you’re seeing some sweet profits roll in. But what if you have multiple Metatrader accounts and want them all to get in on the action? 🤔

In the past, you’d have to buy a separate Forex EA license for each and every account. Not only does that get pricey, but it’s also a lot of extra work to manage all those licenses and accounts. But hold on to your trading hats because there’s a brilliant solution: [Local Trade Copier](https://tools.techidaily.com/mt4copier/)™ software!

If you are looking at how to configure MT4 and MT5 accounts to copy and paste trades from one EA to many other Metatrader accounts, then this is exactly what you need.

**🚀 One License, Unlimited Trading Power 🚀**

With the [Local Trade Copier](https://tools.techidaily.com/mt4copier/)™, you can copy and paste those successful trades from your Forex bot to as many Metatrader accounts as you want, all without buying extra licenses. It’s like having a superpower where one smart trade decision gets multiplied across all your accounts, maximizing your profits without maximizing your expenses!

**🛠️ Easy-Peasy Setup for MT4 and MT5 Accounts 🛠️**

Now, you might be thinking, “This sounds great, but is it complicated to set up?” Nope! In this video, I am going to walk you through the whole process, step by step. I’ll show you how to get your MT4 and MT5 accounts configured and running smoothly with the [Local Trade Copier](https://tools.techidaily.com/mt4copier/)™ and any Forex EA, without any tech headaches. Stop buying Forex EA license for every MT4/MT5 account and watch this video.

## How To Easily Copy & Paste Your Forex and Gold Trades From MT5 to MT4

If you are a Forex trader who uses MetaTrader 5 and MetaTrader 4, you may want to copy your trades from one platform to another. This can be a great way to diversify your trading and take advantage of the unique features of each platform. In this guide, we will show you how to copy your Forex and Gold trades from MT5 to MT4 in one click.

<iframe width="898" height="503" src="https://www.youtube.com/embed/0bT7PHaxQbM" title="How To Easily Copy &amp; Paste Your Forex and Gold Trades From MT5 to MT4 in a Few Simple Steps" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

In this video, Rimantas shows how to easily copy & paste your Forex and Gold trades from MT5 to MT4 in a few simple steps.

Now you can copy Forex and Gold trades from MT5 to MT4 and reap bigger profits in a few simple steps, without fear of making mistakes, even if you’ve never done it before.


Imagine you have a magical notebook. Whatever you write in it, the words magically appear in another notebook too. That’s kind of like what Rimantas is going to show you in this video. It’s like a magic trick for your trades!

Rimantas will teach you how to take your Forex and Gold trades from one book (MT5) and make them appear in another book (MT4) using a special tool called the [Local Trade Copier](https://tools.techidaily.com/mt4copier/) software.

Think of it like copying a drawing from one piece of paper and pasting it onto another without using scissors or glue. It’s that easy!

It’s like having a magic wand that does the work for you. So, even if you’ve never tried this magic trick before, Rimantas will guide you step by step. Get ready to become a trading magician!

## How to Enable Trading in MT4?

![](https://tools.techidaily.com/images/apps/mt4copier/how-to-enable-trading-in-mt4/1.png)

It is disappointing to find the message “Trade disabled” on your MT4 platform. However, that is not fatal, and you can solve it by finding the reason for disabled trading and fixing it. In some cases, if the market is closed, there is nothing to resolve, but if the market is open and your trading functionality is disabled, you have to know what the issue is and how to solve it. In this article, we will look a the main reasons why trading can be disabled on the MetaTrader 4 platform and what to do to enable it.

### What does it mean ‘trade is disabled’?

“Trade is disabled” error message on Metatrader 4 means that you cannot actively trade at all or only some specific instruments, depending on the error message that you get. To find the solution, usually, you need to contact your broker’s support team. But it could be a simple case that the market is closed already for that specific instrument, and you simply need to wait until the market opens.

### Why is my trading disabled on MT4?

There are four main reasons you see an error message, and let’s try to figure out what to do about each one of them.  

**The market is closed.**

Like many other markets such as commodities, stocks, or bonds, Forex is closed during weekends. It means you cannot open a trade because active trading is disabled during weekends. CFD instruments such as  APPL, TSLA, GOLD, etc., are open for trading only during specific market hours (i.e., US stock market hours). So, the solution to the problem is to wait for the markets to open, and you will likely not get the message when they do.

Trading can also be closed if there is a market holiday. You need to check your broker’s calendar to see what days of the year are market holidays.  

Some instruments, usually gold and oil, and some others, may have trading breaks, during which trading functionality is disabled for these instruments. Again, check with your broker to know what instruments have trading breaks and which ones don’t.

You can check trading hours for every trading instrument in the Metatrader 4 Market Watch, but honestly, that information is not correct on most accounts, even with top brokerage firms.

![](https://tools.techidaily.com/images/apps/mt4copier/how-to-enable-trading-in-mt4/2.png)

Trading rarely closed due to a force majeure or unforeseen event. This feature won’t typically last long, but it may happen during times of low liquidity, most often during important news releases, for example, FED interest rates announcements, when markets can become very volatile and move hundreds of pips within a matter of seconds. 

Why is gold trading disabled?

Gold is a commodity, and the commodity market does not open until late hours on Sunday. Depending on the broker you trade, gold opening trading hours will most likely be at 23.00-23.45 EET, depending on the broker you trade. The same can be true for other commodities and some indexes, so be sure to check the opening hours of the instruments of interest with your broker.

**You are logged in with “investor password”**

If you log in to the MT4 account with an investor password, you will only have read-only access. Log in with the primary password to have full trading permissions. It might also happen that your Forex broker support team set your account status to read-only. In that case, you will have to contact the support team of your broker and ask the status to be changed.


**An instrument might be set to “close only” by your broker.**


Occasionally you might find this message on your terminal. It means that your broker is trying to remove that specific instrument from the MT4 platform or any other option of the platforms they offer. Sometimes, as we have stated, some trading instruments become too volatile to trade. Usually, it may be some exotic pairs, for example, Russian ruble or Turkish lira. These have undergone severe geopolitical headwinds in the past. Some brokers decided to remove the currency pairs denominated in the Russian Ruble and Turkish Lira from their MT4 platforms when they became highly volatile. If you have open positions on these pairs, you’ll notice that you can only close them (hence the term “close only”).

**Your account has not yet been activated.**


You might get the message that trading on your MT4 terminal has been disabled because your account has not been activated yet. Usually, when you register with your broker, you will have to go through a verification and activation procedure, which would involve several steps, such as sending some documents or depositing funds before you can start trading. 

After you complete all necessary procedures, the broker will activate your account, and you are good to start trading whenever you want. If you want to check your account status, contact your Account Manager or the broker’s Client Experience team. They will activate your account in case there is any misunderstanding or error.

### Extra possible reason

Suppose you are using automated trading for scalping or similar intra-day trading strategies. “Trading is disabled” errors may cause you to lose profits if you follow a great trader. Well, thighs happen. Such errors can be a weak internet connection, restriction from your broker, or simply malfunctioning of a trading bot.

“Trading is disabled” can also occur if there’s no such financial instrument to open a trade. In that case, you can get the MT4 error code 133, which means your trade copier could not find a currency pair during the execution of your order on the MT4 terminal. Usually, this happens when the client account’s instrument name is different from the master account. For example, the master account sent a trade for EURUSD, but on the client-side, there’s no EURUSD because it is called EURUSDfx. Note the fx suffix.

Not all trade copiers can automatically detect suffixes and adapt. But that is never a problem for our Local Trade Copier.

You might also need to see if you have enabled auto-trading on your MetaTrader terminal. Go to Tools, click on the “Options” menu, click on “Expert Advisor”, and finally allow the automated trading option. Press “OK”, and you are all good.

### Where do I get error messages about disabled trading?

Last but not least, where do you get those error messages about disabled trading on your MT4? You can always check the ‘Journal’ tab at the bottom of the MT4 platform for any error messages when you cannot open or close a trade. There you will usually find an explanation on why you cannot open or close a trade and possibly recommended actions on how to solve the issue. It is also worth checking the ‘Experts’ tab for any error messages or recommendations from the Expert Advisor or trade copier software you are using.

#### How to Enable One Click Trading in MT4?

If you are looking to enable 1-click trading on Metatrader 4, all you have to do is press Alt+T on the keyboard for any particular chart window. That will make the one-click trading buttons appear on the chart, allowing you to quickly open buy and sell trades with a predefined fixed lot size.

![](https://tools.techidaily.com/images/apps/mt4copier/how-to-enable-trading-in-mt4/3.png)

The one-click trading tool on MT4 is very convenient. Still, there’s an even better tool with more functionalities that can also set the lot size automatically after you choose to risk the percentage of your account balance. Check out the Trader On Chart trading panel for MT4, and you’ll never want to open another trade without it.

![](https://tools.techidaily.com/images/apps/mt4copier/how-to-enable-trading-in-mt4/4.png)



## How To Copy Your Forex and Gold Trades From MT4 to MT5 in One Click

If you are a Forex trader who uses MetaTrader 4 and MetaTrader 5, you may want to copy your trades from one platform to another. This can be a great way to diversify your trading and take advantage of the unique features of each platform. In this guide, we will show you how to copy your Forex and Gold trades from MT4 to MT5 in one click.

<iframe width="898" height="503" src="https://www.youtube.com/embed/Xgj2mFUI5qE" title="How To Copy Your Forex and Gold Trades From MT4 to MT5 in One Click" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

In this video, Rimantas shows how to copy your Forex and Gold trades from MT4 to MT5 in one click using the [Local Trade Copier](https://tools.techidaily.com/mt4copier/) software, even if you’ve never linked accounts before.

Have you ever played with walkie-talkies? You press a button, say something, and your friend hears it on their walkie-talkie, even if they’re far away. It’s like magic, right? Well, Rimantas has a similar kind of magic trick to show you in this video, but it’s for your Forex and Gold trades.


Imagine MT4 is your walkie-talkie, and MT5 is your friend’s walkie-talkie. Normally, if you wanted to tell your friend something, you’d have to walk over to them and say it. But with walkie-talkies, you can just press a button, and they hear you! Rimantas is going to teach you how to press that magical button for your trades. With just one click, you can send your Forex and Gold trades from your MT4 walkie-talkie to your friend’s MT5 walkie-talkie. It’s that simple!

Now, let’s think of another fun analogy. Have you ever listened to the radio? Radios are amazing. You turn them on, and you can hear music or people talking from far away places. It’s like they’re broadcasting their voice all over the city, or even multiple cities. That’s what you can do with your trades. You can broadcast them from MT4, just like a radio station sends out songs to many places. And the best part? You can do it in just a few seconds!

You might be thinking, “But I’m not a DJ at a radio station. How can I do this?” Don’t worry! Rimantas is here to help. He’s like the best radio DJ teacher you could ever have. He’ll guide you step by step, showing you how to turn your MT4 into a powerful radio station that can send your trades to MT5. And you don’t need to be a trading expert or a radio DJ to do it. It’s as easy as turning on your favorite song on the radio.

And guess what? You won’t need a bunch of gadgets or devices. It’s not like trying to juggle three balls at once. No, it’s much simpler. Just one click, and your trades fly from MT4 to MT5, just like a song travels through the airwaves to reach radios in many homes.

Maybe you’re a little nervous. Maybe you’re thinking, “What if I make a mistake? What if I’ve never done this before?” It’s okay to feel that way. Remember the first time you tried to ride a bike? It was a bit scary, right? But with someone guiding you, holding the bike steady, you learned. And soon, you were riding all by yourself, feeling the wind in your hair. That’s what Rimantas is here for. He’s like that person holding the bike for you, making sure you don’t fall. He’ll show you everything, even if you’ve never linked accounts before.

So, get ready for an exciting journey. By the end of this video, you’ll be broadcasting your trades like a pro radio DJ, reaching MT5 with ease. You’ll achieve more, with less effort, and have a lot of fun along the way. Tune in, and let Rimantas show you the magic of trading!
<ins class="adsbygoogle"
    style="display:block"
    data-ad-format="autorelaxed"
    data-ad-client="ca-pub-7571918770474297"
    data-ad-slot="1223367746"></ins>

<span class="atpl-alsoreadstyle">Also read:</span>
<div><ul>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-lava-agni-2-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Lava Agni 2 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-to-activate-stellar-data-recovery-for-iphone-6-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Activate Stellar Data Recovery for iPhone 6 | Stellar</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/best-android-data-recovery-undelete-lost-music-from-xiaomi-13t-by-fonelab-android-recover-music/"><u>Best Android Data Recovery - Undelete Lost Music from Xiaomi 13T</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-ispoofer-is-not-working-on-lava-yuva-3-pro-fixed-drfone-by-drfone-virtual-android/"><u>In 2024, iSpoofer is not working On Lava Yuva 3 Pro? Fixed | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-to-bypass-honor-x8b-s-lock-screen-pattern-pin-or-password-by-drfone-android-unlock-android-unlock/"><u>How to bypass Honor X8b’s lock screen pattern, PIN or password</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/can-i-recover-permanently-deleted-photos-from-xiaomi-redmi-note-13-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>Can I recover permanently deleted photos from Xiaomi Redmi Note 13 5G</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/5-ways-to-restart-infinix-note-30i-without-power-button-drfone-by-drfone-reset-android-reset-android/"><u>5 Ways to Restart Infinix Note 30i Without Power Button | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-life360-circle-everything-you-need-to-know-on-vivo-x100-pro-drfone-by-drfone-virtual-android/"><u>In 2024, Life360 Circle Everything You Need to Know On Vivo X100 Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-do-i-sign-a-dot-file-electronically-by-ldigisigner-sign-a-word-sign-a-word/"><u>How do i sign a .dot file electronically</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/c55-unlock-tool-remove-android-phone-password-pin-pattern-and-fingerprint-by-drfone-android-unlock-android-unlock/"><u>C55 Unlock Tool - Remove android phone password, PIN, Pattern and fingerprint</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-life360-circle-everything-you-need-to-know-on-vivo-y78plus-drfone-by-drfone-virtual-android/"><u>In 2024, Life360 Circle Everything You Need to Know On Vivo Y78+ | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-do-you-play-mkv-files-on-redmi-k70-pro-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How do you play MKV files on Redmi K70 Pro?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-best-tools-to-hard-reset-tecno-spark-20-proplus-drfone-by-drfone-reset-android-reset-android/"><u>3 Best Tools to Hard Reset Tecno Spark 20 Pro+ | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/bypassreset-p60-phone-screen-passcodepatternpin-by-drfone-android-unlock-android-unlock/"><u>Bypass/Reset P60 Phone Screen Passcode/Pattern/Pin</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-special-features-virtual-location-on-xiaomi-redmi-a2-drfone-by-drfone-virtual-android/"><u>In 2024, How To Use Special Features - Virtual Location On Xiaomi Redmi A2? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-is-fake-gps-location-spoofer-a-good-choice-on-realme-narzo-60-pro-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Is Fake GPS Location Spoofer a Good Choice On Realme Narzo 60 Pro 5G? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-is-fake-gps-location-spoofer-a-good-choice-on-apple-iphone-x-drfone-by-drfone-virtual-ios/"><u>In 2024, Is Fake GPS Location Spoofer a Good Choice On Apple iPhone X? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/can-t-view-mov-movies-content-on-motorola-moto-g14-by-aiseesoft-video-converter-play-mov-on-android/"><u>Can’t view MOV movies content on Motorola Moto G14</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/can-t-view-hevc-h-265-content-on-redmi-12-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>Can’t view HEVC H.265 content on Redmi 12 5G</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-special-features-virtual-location-on-xiaomi-redmi-a2plus-drfone-by-drfone-virtual-android/"><u>In 2024, How To Use Special Features - Virtual Location On Xiaomi Redmi A2+? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-video-files-on-play-40c-by-fonelab-android-recover-video/"><u>Complete guide for recovering video files on Play 40C</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-music-files-on-oppo-find-x7-ultra-by-fonelab-android-recover-music/"><u>Complete guide for recovering music files on Oppo Find X7 Ultra</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-do-i-repair-and-restore-excel-file-stellar-by-stellar-guide/"><u>How Do I Repair and Restore Excel File? | Stellar</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-music-from-xiaomi-14-by-fonelab-android-recover-music/"><u>Easy steps to recover deleted music from Xiaomi 14</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-contacts-files-on-itel-s23plus-by-fonelab-android-recover-contacts/"><u>Complete guide for recovering contacts files on Itel S23+.</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-pictures-files-on-oppo-k11-5g-by-fonelab-android-recover-pictures/"><u>Complete guide for recovering pictures files on Oppo K11 5G.</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/bypassreset-90-lite-phone-screen-passcodepatternpin-by-drfone-android-unlock-android-unlock/"><u>Bypass/Reset 90 Lite Phone Screen Passcode/Pattern/Pin</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-to-150-2023-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to 150 (2023) Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-pictures-from-samsung-galaxy-a15-4g-by-fonelab-android-recover-pictures/"><u>Easy steps to recover deleted pictures from Samsung Galaxy A15 4G.</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-poco-m6-pro-4g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Poco M6 Pro 4G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-to-bypass-google-frp-on-p55-5g-by-drfone-android-unlock-remove-google-frp/"><u>How To Bypass Google FRP on P55 5G</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/5-ways-to-reset-itel-p40-without-volume-buttons-drfone-by-drfone-reset-android-reset-android/"><u>5 Ways to Reset Itel P40 Without Volume Buttons | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-your-account-has-been-disabled-in-the-app-store-and-itunes-from-apple-iphone-15-by-drfone-ios/"><u>In 2024, Your Account Has Been Disabled in the App Store and iTunes From Apple iPhone 15?</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/unlock-your-poco-m6-pro-5gs-potential-the-top-20-lock-screen-apps-you-need-to-try-by-drfone-android/"><u>Unlock Your Poco M6 Pro 5Gs Potential The Top 20 Lock Screen Apps You Need to Try</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-itel-p55-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Itel P55 5G</u></a></li>
<li><a href="https://techidaily.com/how-to-recover-lost-data-from-apple-iphone-14-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from Apple iPhone 14 Plus? | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/how-can-i-get-more-stardust-in-pokemon-go-on-honor-100-drfone-by-drfone-virtual-android/"><u>How can I get more stardust in pokemon go On Honor 100? | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/sbv-to-srt-how-to-convert-youtube-sbv-subtitle-to-srt-format-for-2024/"><u>SBV to SRT How to Convert YouTube SBV Subtitle to SRT Format for 2024</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/how-can-i-control-speed-of-a-video/"><u>How Can I Control Speed of a Video</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-do-tecno-camon-30-pro-5g-screen-sharing-drfone-by-drfone-android/"><u>How To Do Tecno Camon 30 Pro 5G Screen Sharing | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/a-guide-samsung-galaxy-m54-5g-wireless-and-wired-screen-mirroring-drfone-by-drfone-android/"><u>A Guide Samsung Galaxy M54 5G Wireless and Wired Screen Mirroring | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/best-video-editing-software-for-musicians-and-content-creators-2024/"><u>Best Video Editing Software for Musicians and Content Creators 2024</u></a></li>
<li><a href="https://techidaily.com/how-to-perform-hard-reset-on-honor-90-gt-drfone-by-drfone-reset-android-reset-android/"><u>How to Perform Hard Reset on Honor 90 GT? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-text-messages-from-xiaomi-redmi-a2plus-to-new-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Text Messages from Xiaomi Redmi A2+ to New Phone | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-show-wi-fi-password-on-nubia-red-magic-8s-proplus-by-drfone-android/"><u>How to Show Wi-Fi Password on Nubia Red Magic 8S Pro+</u></a></li>
<li><a href="https://techidaily.com/turn-off-screen-lock-y36-by-drfone-android-unlock-android-unlock/"><u>Turn Off Screen Lock - Y36</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-how-to-unlock-iphone-xr-official-method-to-unlock-your-iphone-xr-by-drfone-ios/"><u>In 2024, How To Unlock iPhone XR Official Method to Unlock Your iPhone XR</u></a></li>
<li><a href="https://techidaily.com/solved-mac-doesnt-recognize-my-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>Solved Mac Doesnt Recognize my iPhone | Stellar</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-2024-approved-20-free-after-effects-logo-reveal-templates/"><u>New 2024 Approved 20 Free After Effects Logo Reveal Templates</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/in-2024-read-and-learn-how-to-convert-a-slow-motion-video-to-normal-in-this-guide-besides-find-the-best-desktop-solution-to-adjust-video-speed-quickly-and-e/"><u>In 2024, Read and Learn How to Convert a Slow-Motion Video to Normal in This Guide. Besides, Find the Best Desktop Solution to Adjust Video Speed Quickly and Easily</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-protecting-your-privacy-how-to-remove-apple-id-from-iphone-14-pro-by-drfone-ios/"><u>In 2024, Protecting Your Privacy How To Remove Apple ID From iPhone 14 Pro</u></a></li>
</ul></div>


