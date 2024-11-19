use anyhow::bail;
use windows::Win32::Globalization::*;
use windows::{core::Result, Win32::System::Threading::*};
use windows::{core::*, Win32::System::Com::*};

static COUNTER: std::sync::RwLock<i32> = std::sync::RwLock::new(0);

extern "system" fn callback(_: PTP_CALLBACK_INSTANCE, _: *mut std::ffi::c_void, _: PTP_WORK) {
    let mut counter = COUNTER.write().unwrap();
    *counter += 1;
}

fn callback_example() -> Result<()> {
    unsafe {
        let work = CreateThreadpoolWork(Some(callback), None, None)?;
        for _ in 0..10 {
            SubmitThreadpoolWork(work);
        }
        WaitForThreadpoolWorkCallbacks(work, false);

        com_example()?;
    }
    println!("counter: {}", COUNTER.read().unwrap());
    Ok(())
}

fn com_example() -> Result<()> {
    unsafe {
        let uri = CreateUri(w!("http://kennykerr.ca"), Uri_CREATE_CANONICALIZE, 0)?;
        let domain = uri.GetDomain()?;
        let port = uri.GetPort()?;
        println!("{domain} ({port})");
    }
    Ok(())
}

struct SpellClient {
    spell_checker: ISpellChecker,
}

impl SpellClient {
    fn try_new(lang: &str) -> anyhow::Result<Self> {
        let spell_checker = unsafe {
            let language_tag = HSTRING::from(lang);
            let spell_checker_factory: ISpellCheckerFactory =
                CoCreateInstance(&SpellCheckerFactory, None, CLSCTX_ALL)?;
            let is_supported = spell_checker_factory.IsSupported(&language_tag)?.as_bool();
            if !is_supported {
                bail!("Language '{lang}' is not supported")
            }
            let spell_checker = spell_checker_factory.CreateSpellChecker(&language_tag)?;
            spell_checker
        };
        Ok(Self { spell_checker })
    }

    fn check(&self, word: &str) -> anyhow::Result<bool> {
        let error = unsafe {
            let text = HSTRING::from(word);
            let spelling_errors = self.spell_checker.Check(&text)?;
            let mut spelling_error = None;
            let result = spelling_errors.Next(&mut spelling_error);
            if result.is_err() {
                bail!("When getting next error: {}", result.message());
            }
            spelling_error
        };
        Ok(error.is_none())
    }

    fn suggest(&self, word: &str) -> anyhow::Result<Vec<String>> {
        let word = HSTRING::from(word);
        unsafe {
            let suggestions = self.spell_checker.Suggest(&word)?;
            todo!();
        }
        Ok(vec![])
    }
}

fn main() -> anyhow::Result<()> {
    let _ = unsafe {
        CoIncrementMTAUsage()?;
    };
    let args: Vec<_> = std::env::args().collect();
    let lang = &args[1];
    let word = &args[2];
    let spell_client = SpellClient::try_new(&lang)?;
    let correct = spell_client.check(word)?;
    if correct {
        println!("No error");
    } else {
        println!("Word is unknown");
        let suggestions = spell_client.suggest(word)?;
        dbg!(suggestions);
    }
    Ok(())
}
