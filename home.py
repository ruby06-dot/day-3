import streamlit as st
st.title("DAY3Contract Formation Advisor")

#1 The ability to accept a hypothetical scenario involving contract formation.
st.subheader("Step 1.  Enter a Hypothetical Scenario")
scenario = st.text_area(
    "Describe the facts",
    placeholder="For example, A offers to sell B a used car for $5,000. B replies by email agreeing to buy it if A provides a roadworthy certificate...",
    height=250,
)

if st.button("Analyse Scenario"):
    if scenario.strip() == "":
        st.warning("Please enter a scenario before analysing.")
    else:
        st.session_state.scenario = scenario
        st.success("Scenario received.")

#The ability to advise on whether a contract has been formed with reference to each of the elements of contract formation.
#The ability to provide a response to the user, on the screen.
#The ability for the user to ask at least one follow-up question to the initial response.
#The ability for the user to download a .docx (Word) version of the response.




